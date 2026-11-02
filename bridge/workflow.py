"""Inbox-zero tools; native reads, explicit filing and bounded search snapshots."""
import datetime as dt
import hashlib
import json
import secrets
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
UTC_RECEIVED = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'

class WorkflowMixin:
    def search(self, query, email=None, limit=50, offset=0, include_deleted=False, cursor=None):
        from outlook_backend import MailError, now, timestamp
        query=str(query).strip(); limit=int(limit); offset=int(offset)
        if not 1<=len(query)<=300 or not 1<=limit<=200 or not 0<=offset<1000:
            raise MailError('Search text required; limit 1..200 and offset 0..999.')
        if not isinstance(include_deleted,bool): raise MailError('include_deleted must be a boolean.')
        session=self.session()
        caches=getattr(self,'searches',{})
        self.searches={k:v for k,v in caches.items() if time.monotonic()-v['created']<300}
        identity=(str(email or '').lower(),query,include_deleted)
        if cursor:
            state=self.searches.get(str(cursor))
            if not state or state['identity']!=identity:
                raise MailError('Search expired or different search parameters. Start again without a cursor.')
        else:
            if email:
                _,binding=self.mailbox(email); bindings={str(email).lower():binding}
            else:
                allowed={a['email'].lower() for a in self.expected}
                bindings={k:v for k,v in self.bindings(session).items() if k in allowed}
            missing=[a['email'] for a in self.expected if not email and a['email'].lower() not in bindings]
            stack=[]; excluded=set()
            for address,binding in bindings.items():
                root=binding[0].GetRootFolder() if binding[2] is None else self.default_folder(binding,6)
                stack.append((address,str(binding[0].StoreID),str(root.EntryID),0))
                if not include_deleted:
                    for number in (3,23):
                        try: excluded.add((str(binding[0].StoreID),str(self.default_folder(binding,number).EntryID)))
                        except Exception: pass
            cursor=secrets.token_hex(16)
            state={'identity':identity,'created':time.monotonic(),'started_at':now(),'stack':stack,
                   'excluded':excluded,'items':[],'total':0,'visited':set(),'scanned':0,'errors':[],
                   'missing':missing,'limited':False}
            if len(self.searches)>=8:
                del self.searches[min(self.searches,key=lambda k:self.searches[k]['created'])]
            self.searches[cursor]=state
        literal=query.replace("'","''")
        fields=['urn:schemas:httpmail:subject','urn:schemas:httpmail:fromemail',
                'http://schemas.microsoft.com/mapi/proptag/0x0C1A001F',
                'http://schemas.microsoft.com/mapi/proptag/0x5D01001F','urn:schemas:httpmail:textdescription']
        restriction='@SQL=("http://schemas.microsoft.com/mapi/proptag/0x001A001F" like \'IPM.Note%\') AND ('+' OR '.join(f'"{f}" like \'%{literal}%\'' for f in fields)+')'
        deadline=time.monotonic()+12
        while state['stack'] and time.monotonic()<deadline:
            address,store_id,folder_id,depth=state['stack'].pop()
            key=(store_id,folder_id)
            if key in state['visited'] or key in state['excluded']:continue
            state['visited'].add(key)
            if depth>25 or len(state['visited'])>2000:
                state['limited']=True;continue
            path=folder_id
            try:
                folder=session.GetFolderFromID(folder_id,store_id)
                if str(folder.StoreID)!=store_id:raise MailError('Search folder does not belong to the mailbox.')
                path=str(folder.FolderPath)
                try:
                    if folder.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003')==2:
                        continue  # Virtual search folders would duplicate physical mail.
                except Exception: pass
                for child in folder.Folders:
                    state['stack'].append((address,store_id,str(child.EntryID),depth+1))
                if folder.DefaultItemType!=0:continue
                table=folder.GetTable(restriction)
                state['scanned']+=1
                count=table.GetRowCount();state['total']+=count
                if not count:continue
                table.Columns.RemoveAll()
                for col in ('EntryID','Subject','SenderName','SenderEmailAddress',UTC_RECEIVED,'UnRead'):
                    table.Columns.Add(col)
                table.Sort('[ReceivedTime]',True)
                # Top 1000 per folder suffice for global top 1000; no body copying.
                for row in table.GetArray(min(count,1000)):
                    identifier,subject,name,sender,received,unread=row
                    state['items'].append({'id':str(identifier),'subject':str(subject or ''),
                        'sender':str(name or ''),'sender_email':str(sender or ''),'received':timestamp(received),
                        'unread':bool(unread),'email':address,'folder_id':folder_id,'folder_path':path})
                state['items'].sort(key=lambda m:(m['received'],m['id']),reverse=True)
                del state['items'][1000:]
            except Exception as exc:
                state['errors'].append({'email':address,'folder':path,'error':str(exc)})
        scanning=bool(state['stack'])
        end=min(offset+limit,len(state['items']))
        return {'items':state['items'][offset:end],'total':state['total'],'offset':offset,
                'next_offset':end if end<len(state['items']) else None,'cursor':cursor,
                'scan_complete':not scanning,'complete':not scanning and not state['errors'] and not state['missing'] and not state['limited'],
                'scanned_folders':state['scanned'],'remaining_branches':len(state['stack']),
                'errors':state['errors'],'missing':state['missing'],'truncated':state['total']>1000 or state['limited'],
                'include_deleted':include_deleted,'started_at':state['started_at'],'read_at':now(),
                'scope':'Locally available mail, including archive and sent subfolders.'}

    def suggest_folder(self,email,item_id):
        from outlook_backend import MailError, now
        item,binding=self.message(email,item_id)
        path=ROOT/'runtime/filing-analysis.json'
        if not path.exists():raise MailError('Folder analysis is missing. Run analyze_filing.py first.')
        analysis=json.loads(path.read_text('utf-8'))
        box=next((b for b in analysis['mailboxes'] if b['email'].lower()==str(email).lower()),None)
        if not box:raise MailError('No folder analysis for this mailbox.')
        sender=str(item.SenderEmailAddress or '').lower(); name=str(item.SenderName or '').lower()
        records=[v for k,v in box['senders'].items() if k==sender or
                 (name and any(str(r['name']).lower()==name for r in v['recent']))]
        candidates={}; total=0
        valid={f['path']:f for f in box['folders'] if f['kind'] not in ('deleted','junk','sent','outbox','draft')}
        for record in records:
            for folder_path,count in record['folders'].items():
                if folder_path in valid:
                    candidates[folder_path]=candidates.get(folder_path,0)+count;total+=count
        result=[]
        for folder_path,count in sorted(candidates.items(),key=lambda r:r[1],reverse=True)[:5]:
            row=valid[folder_path]
            try:
                target=self.folder(self.session(),binding,row['id'])
                # Deleted parent cannot become a positive destination since the analysis.
                parent=target
                forbidden={str(self.default_folder(binding,n).EntryID) for n in (3,23)}
                for _ in range(25):
                    if str(parent.EntryID) in forbidden:raise MailError('Folder is now inside deleted or junk mail.')
                    if str(parent.EntryID)==str(binding[0].GetRootFolder().EntryID):break
                    parent=parent.Parent
                else:raise MailError('Destination folder is too deep to verify its archive path.')
                result.append({'folder_id':str(target.EntryID),'path':str(target.FolderPath),
                    'examples':count,'total_examples':total,'share':round(count/total,3),
                    'confidence':'strong' if count>=5 and count/total>=.9 else 'review',
                    'current_folder':str(item.Parent.EntryID)==str(target.EntryID)})
            except Exception:continue
        return {'email':str(email),'item_id':str(item.EntryID),'subject':str(item.Subject or ''),
                'suggestions':result,'analysis_at':analysis['checked_at'],'read_at':now(),
                'note':'Advice from earlier filing, not an automatic move. Check subject, year and context; form senders can have different destinations.'}

    def move(self,email,item_id,folder_id,source_folder_id=None):
        return self._move(email,item_id,folder_id,source_folder_id)

    def trash(self,email,item_id,source_folder_id,reason):
        from outlook_backend import MailError
        if not str(reason or '').strip():
            raise MailError('Give the reason for cleaning up.')
        _,binding=self.mailbox(email)
        inbox_id=str(self.default_folder(binding,6).EntryID)
        if str(source_folder_id)!=inbox_id:
            raise MailError('Clean-up is only allowed from the Inbox.')
        deleted_id=str(self.default_folder(binding,3).EntryID)
        result=self._move(email,item_id,deleted_id,inbox_id,allow_deleted=True)
        return dict(result,reason=str(reason).strip(),permanent_delete=False)

    def _move(self,email,item_id,folder_id,source_folder_id=None,allow_deleted=False):
        from outlook_backend import MailError, now
        key=hashlib.sha256((str(email).lower()+'\0'+str(item_id)+'\0'+str(folder_id)).encode()).hexdigest()
        journal_path=self.journal_path.with_name('move-journal.json')
        journal=json.loads(journal_path.read_text('utf-8')) if journal_path.exists() else {}
        if key in journal:
            receipt=journal[key]
            if receipt['state']!='moved':raise MailError('Move uncertain. Check source and destination folders before moving again.')
            moved,_=self.message(email,receipt['new_id'])
            if str(moved.Parent.EntryID)!=str(folder_id):raise MailError('This message was filed again after the earlier move. Read the current folder.')
            return dict(receipt,already_moved=True)
        session,binding=self.mailbox(email)
        item,_=self.message(email,item_id)
        target=self.folder(session,binding,folder_id)
        if getattr(target,'DefaultItemType',0)!=0:raise MailError('Choose a mail folder.')
        for number in (4,16):
            if str(item.Parent.EntryID)==str(self.default_folder(binding,number).EntryID):
                raise MailError('Do not move drafts or Outbox mail with this command.')
        if source_folder_id and str(item.Parent.EntryID)!=str(source_folder_id):
            raise MailError('The message has meanwhile moved to a different source folder. Read it again.')
        if str(item.Parent.EntryID)==str(target.EntryID):
            return {'state':'moved','email':str(email),'new_id':str(item.EntryID),'folder':str(target.FolderPath),'already_in_folder':True}
        if allow_deleted and str(target.EntryID)!=str(self.default_folder(binding,3).EntryID):
            raise MailError('Clean-up is only allowed to the own Deleted Items.')
        forbidden={str(self.default_folder(binding,n).EntryID) for n in ((4,16,23) if allow_deleted else (3,4,16,23))}
        parent=target
        for _ in range(25):
            if str(parent.EntryID) in forbidden:raise MailError('This command only archives; do not choose a deleted, junk, draft or outbox folder.')
            if str(parent.EntryID)==str(binding[0].GetRootFolder().EntryID):break
            parent=parent.Parent
        else:raise MailError('Destination folder is too deep to verify its archive path.')
        receipt={'state':'moving','email':str(email),'old_id':str(item.EntryID),'source_folder_id':str(item.Parent.EntryID),
                 'target_folder_id':str(target.EntryID),'subject':str(item.Subject or ''),'unread':bool(item.UnRead),'started_at':now()}
        def persist():
            journal_path.parent.mkdir(exist_ok=True)
            journal[key]=receipt
            temp=journal_path.with_suffix('.tmp');temp.write_text(json.dumps(journal,ensure_ascii=False,indent=2),encoding='utf-8');temp.replace(journal_path)
        persist()
        try:
            moved=item.Move(target)
            if str(moved.Parent.EntryID)!=str(target.EntryID):raise MailError('Outlook did not confirm a correct destination folder.')
            if bool(moved.UnRead)!=receipt['unread']:
                moved.UnRead=receipt['unread'];moved.Save()
            receipt.update(state='moved',new_id=str(moved.EntryID),folder=str(moved.Parent.FolderPath),moved_at=now())
            self.folder_cache={};self.searches={}
            self.invalidate_index('move')
            persist()
            return receipt
        except Exception:
            receipt['state']='uncertain';persist()
            raise MailError('Move outcome uncertain. Check both folders; the operation is not repeated automatically.') from None
