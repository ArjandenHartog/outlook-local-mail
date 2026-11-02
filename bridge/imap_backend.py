"""Direct TLS IMAP/SMTP adapter, independent of Outlook. All IDs bind account + UIDVALIDITY."""
import base64
import contextlib
import datetime as dt
import hashlib
import html
import imaplib
import json
import re
import smtplib
import ssl
import threading
import time
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from email.utils import getaddresses, formataddr, format_datetime, make_msgid
from pathlib import Path
import mail_index
import mail_attachments
import draft_edit

ROOT=Path(__file__).resolve().parent
CONFIG=ROOT/'imap-accounts.json'
IMAP_DB=ROOT/'runtime/imap-index.sqlite3'
IMAP_MARKER=ROOT/'runtime/imap-invalidated.json'
DSN_MARKER='X-Local-Request-DSN'

def accounts():
    if not CONFIG.exists():return []
    return json.loads(CONFIG.read_text('utf-8'))

def utf7_encode(value):
    result=[];pending=[]
    def flush():
        if pending:
            result.append('&'+base64.b64encode(''.join(pending).encode('utf-16be')).decode().rstrip('=').replace('/',',')+'-');pending.clear()
    for char in value:
        if 32<=ord(char)<=126:
            flush();result.append('&-' if char=='&' else char)
        else:pending.append(char)
    flush();return ''.join(result)

def utf7_decode(value):
    def convert(match):
        part=match.group(1)
        if not part:return '&'
        return base64.b64decode(part.replace(',','/')+'='*((-len(part))%4)).decode('utf-16be')
    return re.sub(r'&([^-]*)-',convert,value)

def quote(value):
    return '"'+str(value).replace('\\','\\\\').replace('"','\\"')+'"'

def identifier(email,folder,uid=None,validity=None):
    data={'e':str(email).lower(),'f':folder}
    if uid is not None:data.update(u=int(uid),v=int(validity))
    return ('imap:' if uid is not None else 'imapfolder:')+base64.urlsafe_b64encode(json.dumps(data,separators=(',',':')).encode()).decode().rstrip('=')

def decode(value,email,message=False):
    prefix='imap:' if message else 'imapfolder:'
    try:
        if not str(value).startswith(prefix):raise ValueError()
        payload=str(value)[len(prefix):];result=json.loads(base64.urlsafe_b64decode(payload+'='*((-len(payload))%4)))
        if result['e']!=str(email).lower() or not isinstance(result['f'],str):raise ValueError()
        if message and (int(result['u'])<=0 or int(result['v'])<=0):raise ValueError()
        return result
    except Exception:raise ValueError('Invalid IMAP ID, or an ID from another mailbox.') from None

def check(result):
    kind,rows=result
    if kind!='OK':raise ValueError('The IMAP server did not confirm the command.')
    return rows

def parse_list(row):
    if not isinstance(row,bytes):raise ValueError('Onverwachte IMAP-mapresponse.')
    match=re.fullmatch(rb'\((.*?)\) ("(?:\\.|[^"])*"|NIL) (.*)',row)
    if not match:raise ValueError('Onleesbare IMAP-mapresponse.')
    flags=match[1].decode('ascii').split();delimiter=match[2].decode('ascii').strip('"');name=match[3].decode('ascii')
    if name.startswith('"'):
        if not name.endswith('"'):raise ValueError('Onleesbare IMAP-mapnaam.')
        name=re.sub(r'\\(.)',r'\1',name[1:-1])
    kind=next((kind for flag,kind in (('\\Trash','deleted'),('\\Junk','junk'),('\\Sent','sent'),('\\Drafts','draft')) if flag in flags),'mail')
    return {'wire':name,'name':utf7_decode(name),'delimiter':delimiter,'flags':flags,'kind':kind,'selectable':'\\Noselect' not in flags}

def parse_rows(rows,email,folder,validity):
    result=[]
    for row in rows:
        if not isinstance(row,tuple):continue
        attributes,raw=row;uid=re.search(rb'\bUID (\d+)',attributes);date=re.search(rb'INTERNALDATE "([^"]+)"',attributes)
        if not uid or not date:raise ValueError('Message metadata is missing UID or received date.')
        received=dt.datetime.strptime(date[1].decode('ascii'),'%d-%b-%Y %H:%M:%S %z').astimezone(dt.timezone.utc).isoformat()
        message=BytesParser(policy=policy.default).parsebytes(raw)
        senders=getaddresses([str(message.get('From',''))]);name,address=senders[0] if senders else ('','')
        result.append({'email':email,'store_id':'imap:'+email,'id':identifier(email,folder['wire'],int(uid[1]),validity),
            'folder_id':identifier(email,folder['wire']),'folder_path':'\\\\'+email+'\\'+folder['name'],
            'subject':str(message.get('Subject','')),'sender':name,'sender_email':address,
            'received':received,'unread':b'\\Seen' not in imaplib.ParseFlags(attributes),'kind':folder['kind']})
    return result

class ImapBackend:
    def __init__(self,configs=None,connector=None,smtp_factory=None):
        self.configs={a['email'].lower():a for a in (accounts() if configs is None else configs)}
        self.connector=connector;self.smtp_factory=smtp_factory or smtplib.SMTP_SSL
        self.connections={};self.locks={email:threading.Lock() for email in self.configs};self.status_cache={}
        self.metadata_cache={}

    def secret(self,config):
        import win32crypt
        name=config['credential_file']
        if Path(name).name!=name:raise ValueError('Invalid local credential path.')
        return win32crypt.CryptUnprotectData((ROOT/'runtime'/name).read_bytes(),None,None,None,1)[1].decode('utf-8')

    def config(self,email):
        config=self.configs.get(str(email).lower())
        if not config:raise ValueError('This IMAP mailbox is not allowed.')
        return config

    @contextlib.contextmanager
    def connection(self,email):
        email=str(email).lower();config=self.config(email);lock=self.locks[email]
        if not lock.acquire(timeout=30):raise ValueError('This IMAP mailbox is still busy. Try again later.')
        conn=None
        try:
            cached=self.connections.pop(email,None)
            if cached:
                conn,used=cached
                try:
                    if time.monotonic()-used>60:raise OSError()
                    check(conn.noop())
                except Exception:
                    try:conn.logout()
                    except Exception:pass
                    conn=None
            if conn is None:
                if self.connector:conn=self.connector(config)
                else:
                    conn=imaplib.IMAP4_SSL(config['host'],int(config['imap_port']),ssl_context=ssl.create_default_context(),timeout=20)
                    conn.login(config['username'],self.secret(config))
            yield conn
            self.connections[email]=(conn,time.monotonic());conn=None
        finally:
            if conn:
                try:conn.logout()
                except Exception:pass
            lock.release()

    def list(self,conn):
        folders=[parse_list(row) for row in check(conn.list()) if row]
        for folder in folders:
            ancestors=[f for f in folders if f['kind']!='mail' and folder['wire'].startswith(f['wire']+f['delimiter'])]
            if ancestors:folder['kind']=max(ancestors,key=lambda f:len(f['wire']))['kind']
        return folders

    def select(self,conn,name,readonly=True,validity=None):
        check(conn.select(quote(name),readonly=readonly))
        epoch=int(conn.response('UIDVALIDITY')[1][0])
        if validity is not None and epoch!=int(validity):raise ValueError('This IMAP folder was recreated. Refresh the messages; the old ID is no longer valid.')
        return epoch

    def metadata(self,conn,email,folder,uids=None):
        epoch=self.select(conn,folder['wire'])
        if uids is None:uids=check(conn.uid('search',None,'ALL'))[0].split()
        result=[]
        for start in range(0,len(uids),250):
            selected=b','.join(uids[start:start+250])
            data=check(conn.uid('fetch',selected,'(UID FLAGS INTERNALDATE BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE MESSAGE-ID)])'))
            result.extend(parse_rows(data,email,folder,epoch))
        return result

    def indexed_metadata(self,conn,email,folder,force=False):
        """Reuse headers only when the server proves the entire folder unchanged."""
        key=(email,folder['wire']); fingerprint=None
        capabilities=b' '.join(check(conn.capability())).upper().split()
        if b'CONDSTORE' in capabilities:
            fingerprint=self.folder_fingerprint(conn,folder)
        previous=self.metadata_cache.get(key)
        if not force and fingerprint and previous and previous[0]==fingerprint and previous[1]==folder:
            return [dict(row) for row in previous[2]]
        rows=self.metadata(conn,email,folder)
        if fingerprint:
            after=self.folder_fingerprint(conn,folder)
            if after!=fingerprint:
                self.metadata_cache.pop(key,None)
                raise ValueError('IMAP folder changed while fetching; previous index kept.')
            if len(rows)!=fingerprint[0]:raise ValueError('Not all IMAP headers were received; previous index kept.')
            self.metadata_cache[key]=(fingerprint,dict(folder),[dict(row) for row in rows])
        else:self.metadata_cache.pop(key,None)
        return rows

    @staticmethod
    def folder_fingerprint(conn,folder):
        data=check(conn.status(quote(folder['wire']),'(MESSAGES UIDVALIDITY UIDNEXT UNSEEN HIGHESTMODSEQ)'))[0]
        values=[]
        for field in ('MESSAGES','UIDVALIDITY','UIDNEXT','UNSEEN','HIGHESTMODSEQ'):
            match=re.search(rb'\b'+field.encode()+rb' (\d+)',data)
            if not match:return None
            values.append(int(match[1]))
        # Zero means the server cannot persist modification sequences.
        return tuple(values) if values[-1]>0 else None

    def status(self,email):
        email=str(email).lower();cached=self.status_cache.get(email)
        if cached and time.monotonic()-cached[0]<15:return cached[1]
        config=self.config(email)
        row={'email':email,'connected':False,'inbox_count':None,'latest_received':None,'mode':'IMAP rechtstreeks',
            'protocol':'imap','from_addresses':[email]+config.get('aliases',[]),'error':None}
        try:
            with self.connection(email) as conn:
                data=check(conn.status('INBOX','(MESSAGES UNSEEN)'))[0]
                row.update(connected=True,inbox_count=int(re.search(rb'MESSAGES (\d+)',data)[1]))
        except Exception:row['error']='IMAP connection unavailable. Check the server or the locally stored credentials.'
        self.status_cache[email]=(time.monotonic(),row);return row

    def folders(self,email,refresh=False):
        email=str(email).lower();rows=[];ids={}
        with self.connection(email) as conn:
            for folder in self.list(conn):
                folder_id=identifier(email,folder['wire']);count=unread=0
                if folder['selectable']:
                    data=check(conn.status(quote(folder['wire']),'(MESSAGES UNSEEN)'))[0]
                    count=int(re.search(rb'MESSAGES (\d+)',data)[1]);unread=int(re.search(rb'UNSEEN (\d+)',data)[1])
                rows.append({'id':folder_id,'name':folder['name'].split(folder['delimiter'])[-1],
                    'path':'\\\\'+email+'\\'+folder['name'],'depth':folder['name'].count(folder['delimiter']),
                    'count':count,'unread':unread,'selectable':folder['selectable']})
                if folder['wire'].upper()=='INBOX':ids['inbox_id']=folder_id
                if folder['kind']=='draft':ids['drafts_id']=folder_id
                if folder['kind']=='sent':ids['sent_id']=folder_id
        return dict(folders=rows,**ids,read_at=mail_index.utcnow(),complete=True,cached=False,errors=[],protocol='imap')

    def messages(self,email,folder_id=None,offset=0,limit=50,query=''):
        email=str(email).lower();offset=int(offset);limit=int(limit)
        if offset<0 or not 1<=limit<=200 or len(str(query))>300:raise ValueError('Invalid page size or search text.')
        name=decode(folder_id,email)['f'] if folder_id else 'INBOX'
        with self.connection(email) as conn:
            folder=next((f for f in self.list(conn) if f['wire']==name and f['selectable']),None)
            if not folder:raise ValueError('This IMAP folder does not exist.')
            self.select(conn,name)
            if query:uids=check(conn.uid('search','CHARSET','UTF-8','TEXT',quote(query).encode('utf-8')))[0].split()
            else:uids=check(conn.uid('search',None,'ALL'))[0].split()
            uids.reverse();selected=uids[offset:offset+limit]
            items=self.metadata(conn,email,folder,selected)
        end=offset+len(selected)
        for item in items:item.pop('store_id');item.pop('kind')
        return {'items':items,'total':len(uids),'next_offset':end if end<len(uids) else None,'read_at':mail_index.utcnow(),'complete':True}

    def message(self,conn,email,item_id):
        identity=decode(item_id,email,message=True);self.select(conn,identity['f'],validity=identity['v'])
        rows=check(conn.uid('fetch',str(identity['u']),'(UID FLAGS INTERNALDATE BODY.PEEK[])'))
        tuples=[row for row in rows if isinstance(row,tuple)]
        if len(tuples)!=1:raise ValueError('This IMAP message no longer exists at this location. Refresh the search data.')
        attributes,raw=tuples[0];folder=next((f for f in self.list(conn) if f['wire']==identity['f']),None)
        if not folder:raise ValueError('The message folder has disappeared.')
        summary=parse_rows([(attributes,raw)],email,folder,identity['v'])[0]
        message=BytesParser(policy=policy.default).parsebytes(raw)
        return identity,summary,message,raw,folder

    @staticmethod
    def body(message):
        part=message.get_body(preferencelist=('plain','html'))
        if not part:return ''
        text=part.get_content()
        if part.get_content_type()=='text/html':
            text=re.sub(r'(?is)<(script|style)\b.*?</\1>','',text)
            text=re.sub(r'(?i)<(?:br\s*/?|/p|/div)>','\n',text)
            text=html.unescape(re.sub(r'<[^>]*>','',text))
        return str(text)

    def read(self,email,item_id):
        email=str(email).lower()
        with self.connection(email) as conn:
            identity,summary,message,raw,folder=self.message(conn,email,item_id)
        summary.pop('kind');summary.pop('store_id')
        return dict(summary,to=str(message.get('To','')),cc=str(message.get('Cc','')),bcc=str(message.get('Bcc','')),body=self.body(message),
            is_draft=folder['kind']=='draft',sender_account=email,
            attachments=[{'name':part.get_filename() or '(inline afbeelding)','size':len(part.get_payload(decode=True) or b'')}
                         for part in message.walk() if part.get_filename() or part.get_content_disposition() in ('attachment','inline')])

    def signature_html(self,config):
        """Optional per-account HTML signature: config['signature_file'] (in config dir or runtime) with {{email}}."""
        name=config.get('signature_file')
        if not name:return None
        if Path(name).name!=name:raise ValueError('Invalid signature file name.')
        for folder in (ROOT/'runtime',ROOT):
            if (folder/name).exists():return (folder/name).read_text('utf-8')
        raise ValueError('The configured signature file does not exist: '+name)

    def signatures(self,email):
        config=self.config(email);name=config.get('signature_name','Default') if config.get('signature_file') else None
        return {'email':email,'default_new':name or 'none','default_reply':name or 'none',
            'signatures':[{'name':name,'template':False}] if name else [],
            'from_addresses':[config['email']]+config.get('aliases',[]),'protocol':'imap'}

    def compose(self,config,subject,body,to,cc,from_address,signature,original=None):
        allowed=[config['email'].lower()]+[a.lower() for a in config.get('aliases',[])]
        address=str(from_address or config['email']).lower()
        if address not in allowed:raise ValueError('The sender is not an allowed alias of this mailbox.')
        for value in (subject,to,cc):
            if '\r' in str(value) or '\n' in str(value):raise ValueError('Invalid line break in mail header.')
        footer_html=self.signature_html(config);signature_name=config.get('signature_name','Default')
        if signature not in ((None,signature_name,'','none') if footer_html else (None,'','none')):raise ValueError('This signature does not belong to the IMAP mailbox.')
        message=EmailMessage(policy=policy.SMTP)
        message['From']=formataddr((config.get('display_name',''),address))
        message['To']=str(to);message['Subject']=str(subject)
        if cc:message['Cc']=str(cc)
        message['Date']=format_datetime(dt.datetime.now(dt.timezone.utc))
        message['Message-ID']=make_msgid(domain=address.split('@')[1])
        text=str(body);html_body='<div data-local-authored="1">'+html.escape(text).replace('\n','<br>')+'</div>'
        if footer_html and signature not in ('','none'):
            footer=footer_html.replace('{{email}}',html.escape(address))
            plain_footer=re.sub(r'(?i)<(?:br\s*/?|/p|/div|hr\s*/?)>','\n',footer)
            plain_footer=html.unescape(re.sub(r'<[^>]*>','',plain_footer)).strip()
            text+='\n\n'+plain_footer
            html_body+='<br>'+footer
        if original:
            source_id=str(original.get('Message-ID',''))
            if source_id:
                message['In-Reply-To']=source_id
                message['References']=(str(original.get('References',''))+' '+source_id).strip()
            quoted='\n\nFrom: '+str(original.get('From',''))+'\n'+self.body(original)
            text+=quoted;html_body+='<blockquote>'+html.escape(quoted).replace('\n','<br>')+'</blockquote>'
        message.set_content(text);message.add_alternative(html_body,subtype='html')
        html_part=message.get_payload()[-1]
        def inline(match):
            mime=match[1];data=base64.b64decode(match[2],validate=True);cid='signature-'+hashlib.sha256(data).hexdigest()[:24]
            html_part.add_related(data,maintype='image',subtype=mime.split('/')[1],cid='<'+cid+'>',disposition='inline',filename='signature.png')
            return 'cid:'+cid
        converted=re.sub(r'data:(image/(?:png|jpeg|gif));base64,([A-Za-z0-9+/=]+)',inline,html_body)
        if html_part.is_multipart():html_part.get_payload()[0].set_content(converted,subtype='html')
        return message

    def draft(self,email,subject='',body='',to='',cc='',reply_id=None,signature=None,from_address=None,attachments=None,
              request_delivery_receipt=False,request_read_receipt=False):
        if not isinstance(request_delivery_receipt,bool) or not isinstance(request_read_receipt,bool):
            raise ValueError('request_delivery_receipt and request_read_receipt must be true or false.')
        files=mail_attachments.prepare(attachments)
        email=str(email).lower();config=self.config(email)
        with self.connection(email) as conn:
            folders=self.list(conn);drafts=next((f for f in folders if f['kind']=='draft'),None)
            if not drafts:raise ValueError('No IMAP drafts folder found.')
            original=None
            if reply_id:
                _,_,original,_,_=self.message(conn,email,reply_id)
                subject=subject or 'RE: '+str(original.get('Subject',''))
                to=to or str(original.get('Reply-To') or original.get('From',''))
            message=self.compose(config,subject,body,to,cc,from_address,signature,original)
            for file in files:
                maintype,subtype=file['mime'].split('/',1)
                message.add_attachment(file['data'],maintype=maintype,subtype=subtype,filename=file['name'])
            if request_read_receipt:
                message['Disposition-Notification-To']=str(message['From'])
            if request_delivery_receipt:
                # Private marker; send() removes it and asks SMTP for a DSN instead.
                message[DSN_MARKER]='1'
            raw=message.as_bytes()
            check(conn.append(quote(drafts['wire']),'(\\Draft \\Seen)',None,raw))
            uid_info=conn.response('APPENDUID')[1][0]
            if not uid_info:raise ValueError('Draft append completed, but its ID could not be confirmed. Check Drafts; do not save again automatically.')
            epoch,uid=map(int,uid_info.split());item_id=identifier(email,drafts['wire'],uid,epoch)
        mail_index.invalidate('draft',IMAP_MARKER);self.status_cache.clear()
        return {'id':item_id,'email':email,'sent':False,'subject':subject,'folder':'\\\\'+email+'\\'+drafts['name'],
            'saved_at':mail_index.utcnow(),'sender_email':from_address or email,'signature':(config.get('signature_name','Default') if self.signature_html(config) and signature not in ('','none') else None),
            'attachments':mail_attachments.receipt(files),
            'delivery_receipt_requested':request_delivery_receipt,'read_receipt_requested':request_read_receipt}

    AUTHORED=re.compile(r'\A(\s*<div(?: data-local-authored="1")?>)((?:[^<]|<br>)*)(</div>)')

    @staticmethod
    def _ordinary_attachments(message):
        """(name, parent, part, size) for real attachments; inline signature images are never listed."""
        rows=[]
        def walk(node):
            for child in node.get_payload():
                if child.is_multipart():walk(child)
                elif child.get_content_disposition()=='attachment' and child.get_filename():
                    rows.append((str(child.get_filename()),node,child,len(child.get_payload(decode=True) or b'')))
        if message.is_multipart():walk(message)
        return rows

    def _replace_authored(self,message,body):
        """Swap only the connector-written text; signature, quote and reply headers stay."""
        plain=message.get_body(preferencelist=('plain',));rich=message.get_body(preferencelist=('html',))
        found=None
        if plain is not None and rich is not None:
            source=str(rich.get_content());found=self.AUTHORED.match(source)
        if not found:
            raise ValueError("This draft's own text cannot be delimited (not formatted by the connector). Nothing was changed; edit the text in your mail client.")
        old=html.unescape(found.group(2).replace('<br>','\n'));current=str(plain.get_content()).replace('\r\n','\n')
        if not current.startswith(old):
            raise ValueError('The plain-text and HTML parts of this draft do not match. Nothing was changed.')
        new=str(body).replace('\r\n','\n').replace('\r','\n')
        plain.set_content(new+current[len(old):])
        rich.set_content(source[:found.start(1)]+'<div data-local-authored="1">'+html.escape(new).replace('\n','<br>')+'</div>'+source[found.end(3):],subtype='html')
        for part in (plain,rich):
            if part is not message:del part['MIME-Version']

    def update_draft(self,email,item_id,**fields):
        """Replace a saved draft by an edited copy (APPEND, verify, remove the old UID). Never sends."""
        import delivery
        spec=draft_edit.parse_update(fields);files=mail_attachments.prepare(spec.get('attachments_add'))
        email=str(email).lower();self.config(email)
        path=ROOT/'runtime/imap-draft-journal.json'
        request=hashlib.sha256(json.dumps([email,str(item_id),spec],sort_keys=True,ensure_ascii=False).encode()).hexdigest()
        with self.connection(email) as conn:
            journal=json.loads(path.read_text('utf-8')) if path.exists() else {}
            if request in journal:
                if journal[request]['state']=='updated':return dict(journal[request]['result'],already_updated=True)
                raise ValueError('An earlier change to this draft is uncertain. Check Drafts for duplicate versions; do not retry automatically.')
            sent=ROOT/'runtime/imap-send-journal.json'
            if sent.exists() and delivery.send_key(email,item_id) in json.loads(sent.read_text('utf-8')):
                raise ValueError('This draft was already submitted for sending; it will not be changed.')
            identity,summary,message,raw,folder=self.message(conn,email,item_id)
            if folder['kind']!='draft':raise ValueError('Only a draft in the IMAP drafts folder can be changed.')
            for key,header in (('subject','Subject'),('to','To'),('cc','Cc'),('bcc','Bcc')):
                if key in spec:
                    del message[header]
                    if spec[key]:message[header]=spec[key]
            if 'body' in spec:self._replace_authored(message,spec['body'])
            removals=spec.get('attachments_remove',[]);present=self._ordinary_attachments(message)
            draft_edit.check_removals(removals,present)
            gone={name.lower() for name in removals}
            for _,parent,part,_ in [row for row in present if row[0].lower() in gone]:parent.get_payload().remove(part)
            kept=[row for row in present if row[0].lower() not in gone];names={row[0].lower() for row in kept}
            for file in files:
                if file['name'].lower() in names or sum(f['name'].lower()==file['name'].lower() for f in files)>1:
                    raise ValueError('This draft already has an attachment named '+file['name']+'.')
            draft_edit.check_totals([row[3] for row in kept],[file['size'] for file in files])
            for file in files:
                maintype,subtype=file['mime'].split('/',1)
                message.add_attachment(file['data'],maintype=maintype,subtype=subtype,filename=file['name'])
            data=message.as_bytes(policy=policy.SMTP)
            # Persist intent first so a lost connection never leads to a blind second APPEND.
            journal[request]={'state':'updating','requested_at':mail_index.utcnow()}
            def persist():
                temp=path.with_suffix('.tmp');temp.write_text(json.dumps(journal),encoding='utf-8');temp.replace(path)
            persist()
            try:
                self.select(conn,identity['f'],readonly=False)
                check(conn.append(quote(identity['f']),'(\\Draft \\Seen)',None,data))
                uid_info=conn.response('APPENDUID')[1][0]
                if not uid_info:raise ValueError('No APPENDUID confirmation received.')
                epoch,uid=map(int,uid_info.split());new_id=identifier(email,identity['f'],uid,epoch)
                self.message(conn,email,new_id)  # the replacement must be readable before the old copy goes
            except Exception:
                journal[request]['state']='uncertain';persist()
                raise ValueError('Change uncertain. Check Drafts for an old and a new version; do not save again automatically.') from None
            removed=False;note=None
            try:
                if epoch!=identity['v']:note='UIDVALIDITY of the drafts folder changed; the old draft could not be removed safely. Check Drafts for duplicate versions.'
                elif b'UIDPLUS' not in check(conn.capability())[0].upper().split():note='The server does not support UID EXPUNGE; the old draft was not removed. Check Drafts for duplicate versions.'
                else:
                    self.select(conn,identity['f'],readonly=False,validity=identity['v'])
                    check(conn.uid('store',str(identity['u']),'+FLAGS.SILENT','(\\Deleted)'))
                    check(conn.uid('expunge',str(identity['u'])));removed=True
            except Exception:note='The old draft could not be removed. Check Drafts for duplicate versions.'
            sender=getaddresses([str(message.get('From',''))])
            result={'id':new_id,'previous_id':str(item_id),'id_changed':True,'uid_changed':True,'uidvalidity_changed':epoch!=identity['v'],
                'old_removed':removed,'email':email,'sent':False,'subject':str(message.get('Subject','')),
                'folder':'\\\\'+email+'\\'+folder['name'],'saved_at':mail_index.utcnow(),'changed':sorted(spec),
                'sender_email':sender[0][1] if sender else email,
                'attachments':[{'name':row[0],'size':row[3]} for row in self._ordinary_attachments(message)]}
            if note:result['note']=note
            journal[request]={'state':'updated','result':result};persist()
        mail_index.invalidate('draft',IMAP_MARKER);self.status_cache.clear();return result

    def delete_draft(self,email,item_id,reason,expect_subject=None):
        """Move an unsent draft to the mailbox's own Trash (reversible); only from the Drafts folder."""
        import delivery
        if not str(reason or '').strip():raise ValueError('Give the reason for deleting the draft.')
        email=str(email).lower();self.config(email)
        sent=ROOT/'runtime/imap-send-journal.json'
        if sent.exists() and delivery.send_key(email,item_id) in json.loads(sent.read_text('utf-8')):
            raise ValueError('This draft was already submitted for sending; it will not be deleted.')
        with self.connection(email) as conn:
            identity,summary,message,raw,folder=self.message(conn,email,item_id)
            if folder['kind']!='draft':raise ValueError('Only a draft in the IMAP drafts folder can be deleted.')
            if expect_subject is not None and str(message.get('Subject',''))!=str(expect_subject):
                raise ValueError('The draft subject does not match --expect-subject. Nothing was deleted.')
            trash=next((f for f in self.list(conn) if f['kind']=='deleted' and f['selectable'] and '\\Trash' in f['flags']),None)
            if not trash:raise ValueError('This mailbox has no recognizable trash folder.')
            capabilities=check(conn.capability())[0].upper().split()
            if b'MOVE' not in capabilities or b'UIDPLUS' not in capabilities:raise ValueError('This server does not support a confirmable UID move.')
            self.select(conn,identity['f'],readonly=False,validity=identity['v'])
            try:
                check(conn.uid('move',str(identity['u']),quote(trash['wire'])))
                epoch,old_uid,new_uid=map(int,conn.response('COPYUID')[1][0].split())
                if old_uid!=identity['u']:raise ValueError()
            except Exception:raise ValueError('Deletion uncertain. Check Drafts and trash; do not repeat automatically.') from None
        mail_index.invalidate('draft',IMAP_MARKER);self.status_cache.clear()
        return {'deleted':True,'permanent_delete':False,'email':email,'id':str(item_id),'subject':str(message.get('Subject','')),
            'moved_to':trash['name'],'new_id':identifier(email,trash['wire'],new_uid,epoch),'reason':str(reason).strip(),'deleted_at':mail_index.utcnow()}

    def search(self,query,email=None,limit=50,offset=0,include_deleted=False,cursor=None):
        if not 1<=len(str(query).strip())<=300 or not 1<=int(limit)<=1000 or not 0<=int(offset)<1000:raise ValueError('Ongeldige IMAP-zoekparameters.')
        items=[];errors=[];scanned=0;total=0
        selected=[str(email).lower()] if email else list(self.configs)
        for address in selected:
            with self.connection(address) as conn:
                for folder in self.list(conn):
                    if not folder['selectable'] or not include_deleted and folder['kind'] in ('deleted','junk'):continue
                    self.select(conn,folder['wire']);scanned+=1
                    uids=check(conn.uid('search','CHARSET','UTF-8','TEXT',quote(query).encode('utf-8')))[0].split()
                    total+=len(uids)
                    items.extend(self.metadata(conn,address,folder,uids[-1000:]))
        items.sort(key=lambda m:(m['received'],m['id']),reverse=True)
        offset=int(offset);end=offset+int(limit)
        for item in items:item.pop('store_id');item.pop('kind')
        return {'items':items[offset:end],'total':total,'next_offset':end if end<min(total,1000) else None,
            'read_at':mail_index.utcnow(),'complete':not errors,'scan_complete':True,'cursor':None,
            'errors':errors,'missing':[],'truncated':total>1000,'scanned_folders':scanned}

    def suggest_folder(self,email,item_id):
        message=self.read(email,item_id);suggestions=[]
        try:
            conn=mail_index.read_connection(IMAP_DB)
            try:
                rows=conn.execute("SELECT folder_id,folder_path,count(*) FROM mail WHERE email=? AND sender_email=? AND kind='mail' AND folder_id<>? GROUP BY folder_id,folder_path ORDER BY count(*) DESC",
                    (str(email).lower(),message['sender_email'],identifier(email,'INBOX'))).fetchall()
            finally:conn.close()
            total=sum(row[2] for row in rows)
            with self.connection(email) as imap:
                actual={identifier(email,f['wire']) for f in self.list(imap) if f['kind']=='mail' and f['selectable']}
            for folder_id,path,count in rows[:5]:
                if folder_id not in actual:continue
                suggestions.append({'folder_id':folder_id,'path':path,'examples':count,'total_examples':total,
                    'share':round(count/total,3),'confidence':'strong' if count>=5 and count/total>=.9 else 'review',
                    'current_folder':folder_id==message['folder_id']})
        except (OSError,mail_index.sqlite3.Error):pass
        return {'email':email,'suggestions':suggestions,'note':'Based on earlier archive examples; check subject and relationship.'}

    def send(self,email,item_id):
        email=str(email).lower();config=self.config(email);path=ROOT/'runtime/imap-send-journal.json'
        with self.connection(email) as conn:
            key=hashlib.sha256((email+'\0'+str(item_id)).encode()).hexdigest()
            journal=json.loads(path.read_text('utf-8')) if path.exists() else {}
            if key in journal:
                if journal[key]['state']=='submitted':return journal[key]
                raise ValueError('Earlier send outcome uncertain; check Drafts/Sent. Do not resend automatically.')
            identity,summary,message,raw,folder=self.message(conn,email,item_id)
            if folder['kind']!='draft':raise ValueError('Only a draft from the IMAP drafts folder can be sent.')
            sender=getaddresses([str(message.get('From',''))])
            allowed=[email]+[a.lower() for a in config.get('aliases',[])]
            if len(sender)!=1 or sender[0][1].lower() not in allowed:raise ValueError('The draft has no allowed sender.')
            recipient_headers=[str(message.get(name,'')) for name in ('To','Cc','Bcc') if message.get(name)]
            recipients=[a for _,a in getaddresses(recipient_headers) if a]
            if not recipients or any('@' not in a or '\n' in a or '\r' in a for a in recipients):raise ValueError("Check the draft's recipients.")
            if 'Bcc' in message:del message['Bcc']
            dsn=DSN_MARKER in message
            if dsn:del message[DSN_MARKER]
            raw=message.as_bytes();receipt={'email':email,'sender_email':sender[0][1],'state':'sending','submitted':False,'delivery_verified':False,'requested_at':mail_index.utcnow(),
                'send_ref':key,'fingerprint':{'subject':str(message.get('Subject','')),'message_id':str(message.get('Message-ID','')),
                    'recipients':[{'address':a.lower()} for a in recipients],
                    'attachments':[{'name':part.get_filename(),'size':len(part.get_payload(decode=True) or b'')} for part in message.iter_attachments() if part.get_filename()],
                    'delivery_receipt_requested':dsn,'read_receipt_requested':'Disposition-Notification-To' in message}}
            def persist():
                temp=path.with_suffix('.tmp');temp.write_text(json.dumps(journal),encoding='utf-8');temp.replace(path)
            journal[key]=receipt;persist()
            try:
                with self.smtp_factory(config.get('smtp_host',config['host']),int(config['smtp_port']),context=ssl.create_default_context(),timeout=20) as smtp:
                    smtp.login(config['username'],self.secret(config))
                    options={}
                    if dsn:
                        smtp.ehlo_or_helo_if_needed()
                        receipt['dsn_supported']=bool(smtp.has_extn('dsn'))
                        if receipt['dsn_supported']:
                            options={'mail_options':['RET=HDRS'],'rcpt_options':['NOTIFY=SUCCESS,FAILURE,DELAY']}
                    refused=smtp.sendmail(sender[0][1],recipients,raw,**options)
                receipt.update(state='submitted',submitted=True,submitted_at=mail_index.utcnow(),refused_recipients=list(refused))
            except Exception:
                receipt['state']='uncertain';persist()
                raise ValueError('SMTP outcome uncertain. Check Sent and the recipient; do not resend.') from None
            # Persist successful SMTP submission BEFORE Sent append or draft cleanup.
            persist();receipt['saved_sent']=False;receipt['draft_removed']=False
            try:
                sent=next(f for f in self.list(conn) if f['kind']=='sent')
                check(conn.append(quote(sent['wire']),'(\\Seen)',None,raw));receipt['saved_sent']=True
                self.select(conn,identity['f'],readonly=False,validity=identity['v'])
                capabilities=check(conn.capability())[0].upper().split()
                if b'UIDPLUS' in capabilities:
                    check(conn.uid('store',str(identity['u']),'+FLAGS.SILENT','(\\Deleted)'))
                    check(conn.uid('expunge',str(identity['u'])));receipt['draft_removed']=True
            except Exception:receipt['cleanup_note']='SMTP accepted the mail; check the copy in Sent and the old draft.'
            receipt['note']='Submitted to SMTP. Delivery to the recipient is not confirmed.';persist()
        mail_index.invalidate('send',IMAP_MARKER);self.status_cache.clear();return receipt

    def move(self,email,item_id,folder_id,source_folder_id=None):
        return self._move(email,item_id,folder_id,source_folder_id)

    def trash(self,email,item_id,source_folder_id,reason):
        """Move an INBOX message to this mailbox's own \\Trash folder; never expunge it from there."""
        if not str(reason or '').strip():raise ValueError('Give the reason for cleaning up.')
        email=str(email).lower()
        if decode(source_folder_id,email)['f'].upper()!='INBOX':raise ValueError('Clean-up is only allowed from the Inbox.')
        with self.connection(email) as conn:
            trash=next((f for f in self.list(conn) if f['kind']=='deleted' and f['selectable'] and '\\Trash' in f['flags']),None)
        if not trash:raise ValueError('This mailbox has no recognizable trash folder.')
        result=self._move(email,item_id,identifier(email,trash['wire']),source_folder_id,allow_deleted=True)
        return dict(result,reason=str(reason).strip(),permanent_delete=False)

    def _move(self,email,item_id,folder_id,source_folder_id=None,allow_deleted=False):
        email=str(email).lower();destination=decode(folder_id,email)['f'];identity=decode(item_id,email,message=True)
        if source_folder_id and decode(source_folder_id,email)['f']!=identity['f']:raise ValueError('The source folder has changed.')
        path=ROOT/'runtime/imap-move-journal.json';key=hashlib.sha256((email+'\0'+item_id+'\0'+folder_id).encode()).hexdigest()
        with self.connection(email) as conn:
            journal=json.loads(path.read_text('utf-8')) if path.exists() else {}
            if key in journal:
                previous=journal[key]
                if previous['state']!='moved':raise ValueError('Move uncertain. Check source and destination; do not move again automatically.')
                self.message(conn,email,previous['new_id']);return dict(previous,already_moved=True)
            identity,summary,_,_,source=self.message(conn,email,item_id);folders=self.list(conn)
            target=next((f for f in folders if f['wire']==destination and f['selectable']),None)
            if allow_deleted and (not target or target['kind']!='deleted' or identity['f'].upper()!='INBOX'):raise ValueError('Clean-up is only allowed from the Inbox to the own trash folder.')
            if not target or target['kind'] in (('junk','draft') if allow_deleted else ('deleted','junk','draft')) or source['kind']=='draft':raise ValueError('Use an existing archive folder; draft, delete and junk mutations are not allowed.')
            if destination==identity['f']:return {'state':'moved','email':email,'new_id':item_id,'folder':summary['folder_path'],'already_in_folder':True}
            capabilities=check(conn.capability())[0].upper().split()
            if b'MOVE' not in capabilities or b'UIDPLUS' not in capabilities:raise ValueError('This server does not support a confirmable UID move.')
            self.select(conn,identity['f'],readonly=False,validity=identity['v'])
            receipt={'state':'moving','email':email,'requested_at':mail_index.utcnow(),'unread':summary['unread']}
            journal[key]=receipt
            def persist():
                temp=path.with_suffix('.tmp');temp.write_text(json.dumps(journal),encoding='utf-8');temp.replace(path)
            persist()
            try:
                check(conn.uid('move',str(identity['u']),quote(destination)))
                value=conn.response('COPYUID')[1][0]
                if not value:raise ValueError('No COPYUID confirmation received.')
                epoch,old_uid,new_uid=map(int,value.split())
                if old_uid!=identity['u']:raise ValueError('The server confirmed a different source message.')
                new_id=identifier(email,destination,new_uid,epoch);_,new_summary,_,_,_=self.message(conn,email,new_id)
                if new_summary['unread']!=summary['unread']:raise ValueError('Leesstatus na verplaatsing verschilt.')
                receipt.update(state='moved',new_id=new_id,folder=new_summary['folder_path'],moved_at=mail_index.utcnow());persist()
            except Exception:
                receipt['state']='uncertain';persist();raise ValueError('Move uncertain. Check both folders; do not repeat automatically.') from None
        mail_index.invalidate('move',IMAP_MARKER);self.status_cache.clear();return receipt

    def locate_message(self,email,message_id):
        """Find a Message-ID in an own IMAP mailbox (proof of arrival there)."""
        email=str(email).lower();found=[];scanned=0
        with self.connection(email) as conn:
            for folder in self.list(conn):
                if not folder['selectable'] or folder['kind']=='sent':continue
                epoch=self.select(conn,folder['wire']);scanned+=1
                for uid in check(conn.uid('search',None,'HEADER','Message-ID',quote(message_id)))[0].split():
                    found.append({'id':identifier(email,folder['wire'],int(uid),epoch),'folder_path':folder['name']})
        return {'email':email,'message_id':message_id,'found':bool(found),'items':found,'scanned_folders':scanned,'complete':True,'errors':[]}

    def _report_candidates(self,conn,email,folders,since,deep):
        import delivery
        reports=[];scanned=[]
        for folder in folders:
            if not folder['selectable'] or folder['kind'] in ('sent','draft'):continue
            if not deep and folder['kind'] not in ('deleted','junk') and folder['wire'].upper()!='INBOX':continue
            epoch=self.select(conn,folder['wire']);scanned.append(folder['name'])
            uids=check(conn.uid('search',None,'SINCE',since.strftime('%d-%b-%Y')))[0].split()[-300:]
            if not uids:continue
            rows=check(conn.uid('fetch',b','.join(uids).decode(),'(UID BODY.PEEK[HEADER.FIELDS (FROM SUBJECT CONTENT-TYPE DATE)])'))
            for row in rows:
                if not isinstance(row,tuple):continue
                header=BytesParser(policy=policy.default).parsebytes(row[1],headersonly=True)
                uid=re.search(rb'UID (\d+)',row[0]);ctype=str(header.get('Content-Type','')).lower()
                sender=str(header.get('From','')).lower();subject=str(header.get('Subject',''))
                if not uid or not ('report' in ctype or any(s in sender for s in delivery.REPORT_SENDERS)
                                   or any(s.lower() in subject.lower() for s in delivery.REPORT_SUBJECTS)):continue
                full=[r for r in check(conn.uid('fetch',uid.group(1).decode(),'(BODY.PEEK[])')) if isinstance(r,tuple)]
                if not full:continue
                reports.append({'id':identifier(email,folder['wire'],int(uid.group(1)),epoch),'message_class':'',
                    'subject':subject,'sender':sender,'received':str(header.get('Date','')),'folder_path':folder['name'],
                    'body':full[0][1][:400000].decode('utf-8','replace')})
        return reports,scanned

    def verify_sent(self,email,item_id=None,message_id=None,subject=None,expect_to=None,expect_attachments=None,deep=False):
        import delivery
        if not (item_id or message_id or subject):raise ValueError('Give the draft ID from send, --message-id or --subject.')
        email=str(email).lower();self.config(email);path=ROOT/'runtime/imap-send-journal.json'
        journal=json.loads(path.read_text('utf-8')) if path.exists() else {}
        entry=journal.get(delivery.send_key(email,item_id)) if item_id else None
        expected=dict((entry or {}).get('fingerprint') or {})
        message_id=message_id or expected.get('message_id')
        if item_id and not entry and not message_id:
            with self.connection(email) as conn:
                _,_,original,_,_=self.message(conn,email,item_id)
            message_id=str(original.get('Message-ID',''))
        result={'email':email,'checked_at':mail_index.utcnow(),'source':'IMAP/SMTP',
            'submitted':bool(entry and entry.get('submitted')) or None,'submitted_at':(entry or {}).get('submitted_at'),
            'journal_state':(entry or {}).get('state'),'in_sent_items':False,'outbox_empty':True,'outbox_count':0,'outbox_matching':[],
            'outbox_note':'Direct SMTP has no local Outbox; submitted=true means the SMTP server accepted the mail.',
            'reports':{},'weak_reports':[],'recipient_mailboxes':[]}
        if entry:result.update(dsn_supported=entry.get('dsn_supported'),refused_recipients=entry.get('refused_recipients',[]))
        sent_message=None
        with self.connection(email) as conn:
            folders=self.list(conn);sent=next((f for f in folders if f['kind']=='sent'),None)
            if sent:
                epoch=self.select(conn,sent['wire'])
                if message_id:criteria=(None,'HEADER','Message-ID',quote(message_id))
                else:criteria=('CHARSET','UTF-8','SUBJECT',quote(subject or expected.get('subject','')).encode('utf-8'))
                uids=check(conn.uid('search',*criteria))[0].split()
                result['sent_candidates']=len(uids);result['match_method']='message_id' if message_id else 'subject'
                if uids:
                    uid=int(uids[-1]);rows=check(conn.uid('fetch',str(uid),'(UID FLAGS INTERNALDATE BODY.PEEK[])'))
                    tuples=[row for row in rows if isinstance(row,tuple)]
                    if tuples:
                        sent_message=BytesParser(policy=policy.default).parsebytes(tuples[0][1])
                        result['sent_item']={'id':identifier(email,sent['wire'],uid,epoch),'folder_path':sent['name']}
            if sent_message is not None:
                addresses=[a.lower() for _,a in getaddresses([str(sent_message.get(n,'')) for n in ('To','Cc') if sent_message.get(n)]) if a]
                for row in expected.get('recipients',[]):
                    if row['address'] not in addresses:addresses.append(row['address'])  # Bcc is absent from the Sent copy.
                try:header_date=dt.datetime.strptime(str(sent_message.get('Date')),'%a, %d %b %Y %H:%M:%S %z').astimezone(dt.timezone.utc).isoformat()
                except (TypeError,ValueError):header_date=None
                result.update(in_sent_items=True,submitted=True,message_id=str(sent_message.get('Message-ID','')),subject=str(sent_message.get('Subject','')),
                    sent_at=(entry or {}).get('submitted_at') or header_date,recipients=[{'address':a} for a in addresses],
                    attachments=[{'name':part.get_filename(),'size':len(part.get_payload(decode=True) or b''),'inline':part.get_content_disposition()=='inline'}
                                 for part in sent_message.walk() if part.get_filename()],
                    delivery_receipt_requested=bool((entry or {}).get('dsn_supported')) and bool(expected.get('delivery_receipt_requested')),
                    read_receipt_requested='Disposition-Notification-To' in sent_message)
                result['attachment_count']=len([a for a in result['attachments'] if not a['inline']])
                since=(delivery.parse_time(result['sent_at']) or delivery.utcnow())-dt.timedelta(days=1)
                reports,scanned=self._report_candidates(conn,email,folders,since,bool(deep))
                result['reports'],result['weak_reports']=delivery.evaluate_reports(reports,result,addresses)
                result['bounce_scan']={'since':result['sent_at'],'deep':bool(deep),'folders':scanned,'folders_scanned':len(scanned),
                    'scope':'whole mailbox' if deep else 'INBOX, Junk en Trash','report_candidates':len(reports),'errors':[]}
            else:
                result.update(message_id=message_id,subject=subject or expected.get('subject'),
                    recipients=[{'address':r['address']} for r in expected.get('recipients',[])],attachments=[])
        if expect_to:expected['recipients']=delivery.normalize_addresses(expect_to)
        elif expected.get('recipients'):expected['recipients']=[r['address'] for r in expected['recipients']]
        if subject:expected['subject']=subject
        if expect_attachments is not None:expected['attachment_count']=int(expect_attachments)
        result['expected']={k:v for k,v in expected.items() if k in ('subject','recipients','attachments','attachment_count')}
        result['checks']=delivery.compare_expected(result['expected'],result) if sent_message is not None else {}
        if sent_message is not None and result.get('message_id'):
            for address in [r['address'] for r in result['recipients']]:
                if address in self.configs:
                    located=self.locate_message(address,result['message_id'])
                    result['recipient_mailboxes'].append({'recipient':address,'backend':'imap','checked':True,'found':located['found'],
                        'items':located['items'],'complete':located['complete']})
        return delivery.finalize(result)
