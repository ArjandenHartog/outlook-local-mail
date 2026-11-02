"""Independent, atomic server metadata index; one minute polling, no Outlook dependency."""
import json
import threading
import time
import mail_index
from imap_backend import ImapBackend,IMAP_DB,IMAP_MARKER

STATE=IMAP_DB.with_name('imap-index-state.json')
LOCK=threading.Lock()

def state(**values):
    temp=STATE.with_suffix('.tmp');temp.write_text(json.dumps(dict(values,checked_at=mail_index.utcnow())),encoding='utf-8');temp.replace(STATE)

def refresh(backend,force=False):
    if not backend.configs:return {'requested':False}
    current=mail_index.status(IMAP_DB)
    if not force and current['available'] and not current['stale'] and current['age_seconds']<50:return {'requested':False}
    if not LOCK.acquire(blocking=False):return {'requested':True,'already_building':True}
    writer=None;start=time.monotonic();mark=mail_index.marker(IMAP_MARKER);scanned=0
    try:
        state(state='building');writer=mail_index.SnapshotWriter(IMAP_DB)
        for email in backend.configs:
            with backend.connection(email) as conn:
                for folder in backend.list(conn):
                    if not folder['selectable']:continue
                    writer.add(backend.indexed_metadata(conn,email,folder,force=force));scanned+=1
                    if time.monotonic()-start>180:raise ValueError('IMAP index scan exceeded its time limit.')
        if mail_index.marker(IMAP_MARKER)!=mark:raise ValueError('IMAP mail changed during the scan; previous index kept.')
        meta={'complete':True,'indexed_at':mail_index.utcnow(),'started_at':mail_index.utcnow(),'invalidation':mark,
            'folders':scanned,'mailboxes':sorted(backend.configs),'missing':[],'errors':[],'duration_seconds':round(time.monotonic()-start,3),
            'source':'IMAP-server'}
        writer.publish(meta);state(state='ready',items=writer.count,**meta)
        return {'requested':True,'items':writer.count,'folders':scanned}
    except Exception as exc:
        state(state='failed',error=type(exc).__name__+': '+str(exc),previous_index_preserved=True)
        return {'requested':True,'failed':True,'note':'IMAP index scan failed; previous search data kept.'}
    finally:
        if writer:writer.close()
        LOCK.release()

def run(backend,stop):
    while not stop.is_set():
        refresh(backend)
        stop.wait(60)
