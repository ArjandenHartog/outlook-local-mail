"""Read-only filing and sender statistics across the locally cached mail folders."""
import collections
import datetime as dt
import json
from pathlib import Path
import pythoncom
import win32com.client

def analyze():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject('Outlook.Application')
    (Path(__file__).resolve().parent/'runtime').mkdir(exist_ok=True)
    report = {'checked_at':dt.datetime.now(dt.timezone.utc).isoformat(),
              'scope':'local cached Outlook mail only', 'mailboxes':[], 'errors':[]}
    for account in app.Session.Accounts:
        store = account.DeliveryStore
        defaults = {}
        for number, kind in ((3,'deleted'),(4,'outbox'),(5,'sent'),(6,'inbox'),(16,'draft'),(23,'junk')):
            try: defaults[store.GetDefaultFolder(number).EntryID] = kind
            except Exception: pass
        box = {'email':account.SmtpAddress,'folders':[], 'senders':{},'mail_count':0}
        senders = collections.defaultdict(lambda: {'count':0,'folders':collections.Counter(),'recent':[]})
        def walk(folder, inherited='archive'):
            kind = defaults.get(folder.EntryID, inherited)
            if folder.DefaultItemType == 0:
                try:
                    table=folder.GetTable(); table.Columns.RemoveAll()
                    for col in ('EntryID','Subject','SenderName','SenderEmailAddress','ReceivedTime','MessageClass'):
                        table.Columns.Add(col)
                    count=0
                    while not table.EndOfTable:
                        for row in table.GetArray(500):
                            entry, subject, name, sender, received, msgclass = row
                            if not str(msgclass).startswith('IPM.Note'): continue
                            count+=1
                            if kind in ('sent','outbox','draft'): continue
                            sender=str(sender or name or 'unknown').lower()
                            s=senders[sender]; s['count']+=1; s['folders'][folder.FolderPath]+=1
                            stamp=str(received)
                            s['recent'].append({'id':entry,'subject':str(subject),'name':str(name),
                                'received':stamp,'folder':folder.FolderPath,'kind':kind})
                            s['recent'].sort(key=lambda x:x['received'],reverse=True)
                            del s['recent'][6:]
                    box['folders'].append({'id':folder.EntryID,'path':folder.FolderPath,'kind':kind,'mail_count':count})
                    box['mail_count']+=count
                except Exception as exc:
                    report['errors'].append({'email':account.SmtpAddress,'folder':folder.FolderPath,'error':str(exc)})
            for child in folder.Folders: walk(child, kind)
        walk(store.GetRootFolder())
        for sender,data in senders.items():
            data['folders']=dict(data['folders']); box['senders'][sender]=data
        report['mailboxes'].append(box)
    (Path(__file__).resolve().parent/'runtime/filing-analysis.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    for b in report['mailboxes']:
        print(json.dumps({'email':b['email'],'folders':len(b['folders']),'cached_mails':b['mail_count'],
            'largest_folders':sorted(b['folders'],key=lambda f:f['mail_count'],reverse=True)[:7],
            'top_senders':[{'sender':k,'count':v['count'],'folders':v['folders'],
                            'subjects':[m['subject'] for m in v['recent'][:3]]}
                for k,v in sorted(b['senders'].items(),key=lambda t:t[1]['count'],reverse=True)[:5]]},ensure_ascii=False))
    print(json.dumps({'errors':report['errors']},ensure_ascii=False))
    pythoncom.CoUninitialize()

if __name__=='__main__': analyze()
