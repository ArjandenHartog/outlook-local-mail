"""Console-free read-only Outlook metadata scanner; attaches to existing Outlook."""
import json
import os
import time
from pathlib import Path
import pythoncom
import win32api
import win32com.client
import win32event
import mail_index
from outlook_backend import OutlookBackend, timestamp
from workflow import UTC_RECEIVED

ROOT=Path(__file__).resolve().parent
SMTP='http://schemas.microsoft.com/mapi/proptag/0x5D01001F'

class ScanFailure(RuntimeError):
    def __init__(self,message,**details):
        super().__init__(message);self.details=details

def progress(**values):
    temp=mail_index.BUILD_STATUS.with_suffix('.tmp')
    temp.write_text(json.dumps(dict(values,pid=os.getpid(),checked_at=mail_index.utcnow())),encoding='utf-8')
    temp.replace(mail_index.BUILD_STATUS)

def scan():
    start=time.monotonic();invalidation=mail_index.marker();indexed_start=mail_index.utcnow()
    # No Dispatch/cold start and no login; the service manages Outlook readiness.
    client=win32com.client.GetActiveObject('Outlook.Application')
    backend=OutlookBackend(ROOT/'accounts.json',client=client)
    session=backend.session();allowed={a['email'].lower() for a in backend.expected}
    bindings={k:v for k,v in backend.bindings(session).items() if k in allowed}
    missing=sorted(allowed-set(bindings));errors=[];folders=0;visited=set();stack=[];roles={}
    for email,binding in bindings.items():
        store_id=str(binding[0].StoreID)
        for number,kind in ((3,'deleted'),(23,'junk'),(16,'draft'),(4,'outbox'),(5,'sent')):
            try:roles[(store_id,str(backend.default_folder(binding,number).EntryID))]=kind
            except Exception as exc:errors.append({'email':email,'default_folder':number,'error':str(exc)})
        root=binding[0].GetRootFolder() if binding[2] is None else backend.default_folder(binding,6)
        stack.append((email,store_id,str(root.EntryID),0,'mail'))
    writer=mail_index.SnapshotWriter()
    try:
        while stack:
            email,store_id,folder_id,depth,kind=stack.pop();key=(store_id,folder_id)
            if key in visited:continue
            visited.add(key)
            if depth>25 or len(visited)>2000 or time.monotonic()-start>240:
                raise RuntimeError('Index scan exceeded the safe limit; previous index kept.')
            folder=session.GetFolderFromID(folder_id,store_id)
            path=str(folder.FolderPath)
            try:
                if folder.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003')==2:continue
            except Exception:pass
            kind=roles.get(key,kind)
            try:
                for child in folder.Folders:stack.append((email,store_id,str(child.EntryID),depth+1,kind))
                if folder.DefaultItemType!=0:continue
                table=folder.GetTable('@SQL="http://schemas.microsoft.com/mapi/proptag/0x001A001F" like \'IPM.Note%\'')
                table.Columns.RemoveAll()
                for column in ('EntryID','Subject','SenderName','SenderEmailAddress',SMTP,UTC_RECEIVED,'UnRead'):
                    table.Columns.Add(column)
                while not table.EndOfTable:
                    rows=[]
                    for identifier,subject,name,raw,smtp,received,unread in table.GetArray(500):
                        sender=smtp if isinstance(smtp,str) and '@' in smtp else raw
                        rows.append(dict(email=email,store_id=store_id,id=str(identifier),folder_id=folder_id,
                            folder_path=path,subject=str(subject or ''),sender=str(name or ''),
                            sender_email=str(sender or ''),received=timestamp(received),unread=int(bool(unread)),kind=kind))
                    writer.add(rows);pythoncom.PumpWaitingMessages()
                folders+=1
                if folders%20==0:progress(state='building',folders=folders,items=writer.count,started_at=indexed_start)
            except Exception as exc:errors.append({'email':email,'folder':path,'error':str(exc)})
        if errors or missing:raise ScanFailure('Not all mailboxes or folders are readable; previous index kept.',errors=errors,missing=missing,folders=folders)
        if mail_index.marker()!=invalidation:raise RuntimeError('Mail changed during the scan; previous index kept, try again later.')
        meta={'complete':True,'indexed_at':mail_index.utcnow(),'started_at':indexed_start,'invalidation':invalidation,
              'folders':folders,'mailboxes':sorted(bindings),'missing':missing,'errors':errors,
              'duration_seconds':round(time.monotonic()-start,3)}
        writer.publish(meta);progress(state='ready',items=writer.count,**meta)
        return dict(meta,items=writer.count)
    finally:writer.close()

def main():
    mutex=win32event.CreateMutex(None,False,'Local\\OutlookLocalMailMetadataIndex')
    acquired=win32event.WaitForSingleObject(mutex,0)
    if acquired not in (win32event.WAIT_OBJECT_0,win32event.WAIT_ABANDONED):return
    pythoncom.CoInitialize()
    try:
        progress(state='building',started_at=mail_index.utcnow(),folders=0,items=0)
        scan()
    except Exception as exc:progress(state='failed',error=str(exc),previous_index_preserved=True,**getattr(exc,'details',{}))
    finally:
        pythoncom.CoUninitialize();win32event.ReleaseMutex(mutex);win32api.CloseHandle(mutex)

if __name__=='__main__':main()
