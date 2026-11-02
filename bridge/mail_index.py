"""Private SQLite metadata snapshots. No bodies, attachments or authentication data."""
import datetime as dt
import json
import re
import sqlite3
import time
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DB = ROOT / 'runtime/mail-index.sqlite3'
INVALIDATION = ROOT / 'runtime/index-invalidated.json'
BUILD_STATUS = ROOT / 'runtime/index-build.json'
FIELDS = ('email','store_id','id','folder_id','folder_path','subject','sender','sender_email','received','unread','kind')

def utcnow():
    return dt.datetime.now(dt.timezone.utc).isoformat()

def marker(path=INVALIDATION):
    try: return Path(path).read_text('utf-8')
    except FileNotFoundError: return ''

def invalidate(reason,path=INVALIDATION):
    path=Path(path)
    path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps({'id':uuid.uuid4().hex,'at':utcnow(),'reason':reason}),encoding='utf-8')
    temp.replace(path)

def read_connection(path):
    return sqlite3.connect(Path(path).resolve().as_uri()+'?mode=ro',uri=True,timeout=2)

class SnapshotWriter:
    """Build privately beside the database, publish only a complete snapshot."""
    def __init__(self,path=DB):
        self.path=Path(path)
        self.path.parent.mkdir(exist_ok=True)
        self.temp=self.path.with_name(self.path.name+'.building-'+uuid.uuid4().hex)
        self.connection=sqlite3.connect(self.temp)
        self.connection.executescript('''
            CREATE TABLE mail(email TEXT, store_id TEXT, id TEXT, folder_id TEXT,
                folder_path TEXT, subject TEXT, sender TEXT, sender_email TEXT,
                received TEXT, unread INTEGER, kind TEXT, UNIQUE(email,id));
            CREATE INDEX mail_scope ON mail(email,folder_id,received DESC);
            CREATE INDEX mail_date ON mail(received DESC);
            CREATE VIRTUAL TABLE words USING fts5(subject,sender,sender_email,folder_path,
                tokenize='unicode61 remove_diacritics 2');
            CREATE TABLE meta(value TEXT NOT NULL);
        ''')
        self.count=0

    def add(self,rows):
        for row in rows:
            # Preserve rowid alignment; duplicate physical identities are errors.
            values=[row[field] for field in FIELDS]
            cursor=self.connection.execute('INSERT INTO mail VALUES (?,?,?,?,?,?,?,?,?,?,?)',values)
            self.connection.execute('INSERT INTO words(rowid,subject,sender,sender_email,folder_path) VALUES (?,?,?,?,?)',
                (cursor.lastrowid,row['subject'],row['sender'],row['sender_email'],row['folder_path']))
            self.count+=1

    def publish(self,meta):
        if not meta.get('complete'):raise ValueError('An incomplete index is not published.')
        self.connection.execute('INSERT INTO meta VALUES (?)',(json.dumps(dict(meta,items=self.count)),))
        self.connection.execute("INSERT INTO words(words) VALUES ('optimize')")
        self.connection.commit();self.connection.close();self.connection=None
        # All readers close promptly; retry short Windows sharing conflicts.
        for attempt in range(10):
            try:self.temp.replace(self.path);return
            except PermissionError:
                if attempt==9:raise
                time.sleep(.1)

    def close(self):
        if self.connection:self.connection.close();self.connection=None
        self.temp.unlink(missing_ok=True)

def status(path=DB):
    is_imap=Path(path).name=='imap-index.sqlite3'
    result={'available':False,'complete':False,'stale':True,'scope':'Local metadata: subject, sender and folder path; no message text.'}
    try:
        connection=read_connection(path)
        try:meta=json.loads(connection.execute('SELECT value FROM meta').fetchone()[0])
        finally:connection.close()
        age=max(0,(dt.datetime.now(dt.timezone.utc)-dt.datetime.fromisoformat(meta['indexed_at'])).total_seconds())
        current_marker=marker(Path(path).with_name('imap-invalidated.json')) if is_imap else marker()
        result.update(meta,available=True,age_seconds=round(age,1),stale=age>(180 if is_imap else 600) or meta.get('invalidation','')!=current_marker)
    except (sqlite3.Error,FileNotFoundError,OSError,ValueError,TypeError):pass
    build_path=Path(path).with_name('imap-index-state.json') if is_imap else BUILD_STATUS
    try:result['build']=json.loads(build_path.read_text('utf-8'))
    except (FileNotFoundError,ValueError):pass
    return result

def allowed_accounts():
    from imap_backend import accounts
    return {a['email'].lower() for a in json.loads((ROOT/'accounts.json').read_text('utf-8-sig'))}|{a['email'].lower() for a in accounts()}

def quick_search(query='',email=None,limit=50,offset=0,include_deleted=False,sender='',folder_id=None,
                 after=None,before=None,unread=None,path=DB,allowed=None,max_limit=200):
    started=time.perf_counter()
    allowed=allowed_accounts() if allowed is None else {a.lower() for a in allowed}
    if email and str(email).lower() not in allowed:raise ValueError('This mailbox is not allowed.')
    limit=int(limit);offset=int(offset);query=str(query).strip();sender=str(sender).strip()
    if not 1<=limit<=max_limit or not 0<=offset<=100000 or len(query)>300 or len(sender)>300:
        raise ValueError('Invalid search text, page size or offset.')
    if not isinstance(include_deleted,bool) or unread is not None and not isinstance(unread,bool):
        raise ValueError('Filters must be booleans.')
    if folder_id and not email:raise ValueError('Choose a mailbox when using a folder filter.')
    state=status(path)
    if not state['available']:raise ValueError('The local search index is being built. Use full-text search for now or try again later.')
    where=['m.email IN ('+','.join('?' for _ in allowed)+')'];params=list(sorted(allowed))
    tokens=re.findall(r'\w+',query,flags=re.UNICODE)
    if query and not tokens:raise ValueError('Search with at least one letter or digit.')
    if tokens:
        where.append('m.rowid IN (SELECT rowid FROM words WHERE words MATCH ?)')
        params.append(' AND '.join('"'+token+'"*' for token in tokens))
    if email:where.append('m.email=?');params.append(str(email).lower())
    if folder_id:where.append('m.folder_id=?');params.append(str(folder_id))
    if not include_deleted:where.append("m.kind NOT IN ('deleted','junk')")
    if sender:
        literal=sender.replace('\\','\\\\').replace('%','\\%').replace('_','\\_')
        where.append("(m.sender LIKE ? ESCAPE '\\' OR m.sender_email LIKE ? ESCAPE '\\')")
        params.extend(['%'+literal+'%']*2)
    dates={}
    for name,value,operator in (('after',after,'>='),('before',before,'<')):
        if value:
            date=dt.date.fromisoformat(str(value));dates[name]=date
            where.append('m.received'+operator+'?');params.append(date.isoformat()+'T00:00:00+00:00')
    if len(dates)==2 and dates['after']>=dates['before']:raise ValueError('The before date must be after the after date.')
    if unread is not None:where.append('m.unread=?');params.append(int(unread))
    sql=' FROM mail m WHERE '+' AND '.join(where)
    connection=read_connection(path)
    try:
        total=connection.execute('SELECT count(*)'+sql,params).fetchone()[0]
        connection.row_factory=sqlite3.Row
        items=[dict(row) for row in connection.execute('SELECT m.*'+sql+' ORDER BY m.received DESC,m.email,m.id LIMIT ? OFFSET ?',params+[limit,offset])]
    finally:connection.close()
    for item in items:item.pop('store_id');item.pop('kind');item['unread']=bool(item['unread'])
    end=offset+len(items)
    return {'items':items,'total':total,'offset':offset,'next_offset':end if end<total else None,
            'read_at':utcnow(),'indexed_at':state['indexed_at'],'age_seconds':state['age_seconds'],
            'stale':state['stale'],'complete':state['complete'] and not state['stale'],
            'missing':state.get('missing',[]),'errors':state.get('errors',[]),'scope':state['scope'],
            'query_seconds':round(time.perf_counter()-started,4),'metadata_only':True}

def sources(email=None):
    from imap_backend import accounts,IMAP_DB
    direct={a['email'].lower() for a in accounts()}
    outlook={a['email'].lower() for a in json.loads((ROOT/'accounts.json').read_text('utf-8-sig'))}
    if email:
        email=str(email).lower()
        if email in direct:return [(IMAP_DB,{email})]
        if email in outlook:return [(DB,{email})]
        raise ValueError('This mailbox is not allowed.')
    return [(DB,outlook)]+([(IMAP_DB,direct)] if direct else [])

def combined_status():
    providers=[dict(status(path),provider='IMAP' if path.name=='imap-index.sqlite3' else 'Outlook') for path,_ in sources()]
    return {'available':any(p['available'] for p in providers),'complete':all(p['available'] and p['complete'] for p in providers),
        'stale':any(p['stale'] for p in providers),'items':sum(p.get('items',0) for p in providers),
        'folders':sum(p.get('folders',0) for p in providers),'sources':providers,
        'indexed_at':min((p['indexed_at'] for p in providers if p.get('indexed_at')),default=None),
        'read_at':utcnow(),'scope':'Local metadata: Outlook cache plus direct IMAP; no message text.'}

def combined_search(**args):
    started=time.perf_counter();limit=int(args.get('limit',50));offset=int(args.get('offset',0))
    if not 1<=limit<=200 or not 0<=offset<=100000:raise ValueError('Invalid page size or offset.')
    items=[];total=0;providers=[];missing=[];errors=[]
    for path,allowed in sources(args.get('email')):
        current=status(path)
        providers.append(dict(current,provider='IMAP' if path.name=='imap-index.sqlite3' else 'Outlook'))
        if not current['available']:
            missing.extend(sorted(allowed));continue
        params=dict(args,limit=offset+limit,offset=0,path=path,allowed=allowed,max_limit=100200)
        result=quick_search(**params);items.extend(result['items']);total+=result['total'];errors.extend(result['errors'])
    if not any(p['available'] for p in providers):raise ValueError('The search index is not available yet. Check index-status.')
    items.sort(key=lambda m:(m['received'],m['email'],m['id']),reverse=True);end=offset+limit
    dates=[p['indexed_at'] for p in providers if p.get('indexed_at')]
    return {'items':items[offset:end],'total':total,'offset':offset,'next_offset':end if end<total else None,
        'read_at':utcnow(),'indexed_at':min(dates),'age_seconds':max(p.get('age_seconds',0) for p in providers),
        'complete':all(p['available'] and p['complete'] and not p['stale'] for p in providers),
        'stale':any(p['stale'] for p in providers),'missing':missing,'errors':errors,'sources':providers,
        'metadata_only':True,'query_seconds':round(time.perf_counter()-started,4),'scope':'Subject, sender and folder path.'}
