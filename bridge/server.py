"""Loopback-only local API with serialized Outlook COM calls; no cloud service."""
import concurrent.futures
import hmac
import json
import os
from pathlib import Path
import queue
import secrets
import socket
import threading
import time
import mail_index
from imap_backend import ImapBackend
import index_imap
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit

from outlook_backend import OutlookBackend, MailError

ROOT = Path(__file__).resolve().parent
RUNTIME = ROOT / 'runtime'
PORT = 38473
ORIGIN = f'http://127.0.0.1:{PORT}'


class Worker:
    def __init__(self):
        self.queue = queue.Queue(maxsize=100)
        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        backend = OutlookBackend(ROOT / 'accounts.json')
        next_connection_check = 0
        try:
            while True:
                if time.monotonic() >= next_connection_check:
                    # Keep COM alive without opening windows, reading message bodies,
                    # or invoking send/receive groups. Recover after Outlook closes.
                    try:
                        session = backend.session()
                        connection = {'connected': True, 'accounts': session.Accounts.Count}
                    except Exception:
                        connection = {'connected': False, 'accounts': 0}
                    connection.update(checked_at=time.time(), pid=os.getpid())
                    if RUNTIME.is_dir():
                        state_path = RUNTIME / 'background.json'
                        temp_path = state_path.with_suffix('.tmp')
                        temp_path.write_text(json.dumps(connection), encoding='utf-8')
                        temp_path.replace(state_path)
                    next_connection_check = time.monotonic() + 60
                try:
                    future, action, args = self.queue.get(timeout=0.2)
                except queue.Empty:
                    pythoncom.PumpWaitingMessages()
                    continue
                if future.set_running_or_notify_cancel():
                    try:
                        future.set_result(getattr(backend, action)(**args))
                    except Exception as exc:
                        future.set_exception(MailError(str(exc)))
                pythoncom.PumpWaitingMessages()
        finally:
            pythoncom.CoUninitialize()

    def submit(self, action, args):
        future = concurrent.futures.Future()
        try:
            self.queue.put_nowait((future, action, args))
        except queue.Full:
            raise MailError('Outlook is busy. Wait until the current commands have finished.') from None
        return future


class Manager(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = False

    def server_bind(self):
        # Windows SO_REUSEADDR permits two listeners, corrupting token discovery.
        if hasattr(socket, 'SO_EXCLUSIVEADDRUSE'):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()

    def __init__(self, address, backend_worker=None):
        super().__init__(address, Handler)
        self.origin = f'http://127.0.0.1:{self.server_port}'
        self.token = secrets.token_urlsafe(32)
        self.worker = backend_worker or Worker()
        self.jobs = {}
        self.jobs_lock = threading.Lock()
        self.index_pool = concurrent.futures.ThreadPoolExecutor(max_workers=2,thread_name_prefix='metadata')
        self.imap = ImapBackend(configs=[] if backend_worker else None)
        self.imap_pool = concurrent.futures.ThreadPoolExecutor(max_workers=2,thread_name_prefix='imap')
        self.router_pool = concurrent.futures.ThreadPoolExecutor(max_workers=4,thread_name_prefix='mail-router')
        self.index_stop = threading.Event()
        if self.imap.configs:
            threading.Thread(target=index_imap.run,args=(self.imap,self.index_stop),daemon=True,name='imap-index').start()

    def unified_status(self):
        result=self.worker.submit('status',{}).result(timeout=120)
        for email in self.imap.configs:
            row=self.imap.status(email);result['accounts'].append(row)
            result['expected']+=1;result['connected']+=int(row['connected'])
        result['source']='Outlook Classic + rechtstreekse IMAP'
        result['imap_refresh_seconds']=60
        return result

    def full_search(self,args):
        limit=int(args.get('limit',50));offset=int(args.get('offset',0))
        if not 1<=limit<=200 or not 0<=offset<1000:raise ValueError('Invalid search page.')
        params=dict(args,limit=200,offset=0)
        try:
            native=self.worker.submit('search',params).result(timeout=240)
            items=list(native['items'])
            if native['scan_complete']:
                for page in range(200,min(offset+limit,1000),200):
                    params.update(cursor=native['cursor'],offset=page)
                    extra=self.worker.submit('search',params).result(timeout=30);items.extend(extra['items'])
        except Exception as exc:
            native={'items':[],'total':0,'complete':False,'scan_complete':True,'errors':[{'source':'Outlook','error':str(exc)}],
                    'missing':[],'scanned_folders':0,'cursor':None};items=[]
        try:
            direct=self.imap.search(query=args['query'],include_deleted=args.get('include_deleted',False),limit=1000)
            items.extend(direct['items'])
        except Exception:
            direct={'total':0,'complete':False,'errors':[{'source':'IMAP','error':'Direct IMAP search failed.'}],
                    'missing':list(self.imap.configs),'scanned_folders':0}
        items.sort(key=lambda m:(m['received'],m['id']),reverse=True)
        total=native['total']+direct['total'];end=offset+limit
        result=dict(native,items=items[offset:end],total=total,next_offset=end if end<min(total,1000) else None,
            complete=native['complete'] and direct['complete'],errors=native['errors']+direct['errors'],
            missing=native['missing']+direct['missing'],scanned_folders=native['scanned_folders']+direct['scanned_folders'],
            truncated=total>1000 or native.get('truncated',False),read_at=mail_index.utcnow())
        return result

    def verify_sent(self,args):
        import delivery
        email=str(args.get('email','')).lower()
        if email in self.imap.configs:result=self.imap.verify_sent(**args)
        else:result=self.worker.submit('verify_sent',args).result(timeout=240)
        message_id=result.get('message_id')
        if result.get('in_sent_items') and message_id:
            # Each backend checks recipients it owns; cross-check the other one here.
            outlook=mail_index.allowed_accounts()-set(self.imap.configs)
            checked={row['recipient'] for row in result.get('recipient_mailboxes',[])}
            for address in [r.get('address') for r in result.get('recipients',[])]:
                if not address or address in checked:continue
                backend='imap' if address in self.imap.configs else 'outlook' if address in outlook else None
                if not backend:continue
                try:
                    if backend=='imap':located=self.imap.locate_message(address,message_id)
                    else:located=self.worker.submit('locate_message',{'email':address,'message_id':message_id}).result(timeout=120)
                    result['recipient_mailboxes'].append({'recipient':address,'backend':backend,'checked':True,
                        'found':located['found'],'items':located['items'],'complete':located['complete']})
                except Exception as exc:
                    result['recipient_mailboxes'].append({'recipient':address,'backend':backend,'checked':False,'found':False,'error':str(exc)[:200]})
            delivery.finalize(result)
            if email not in self.imap.configs and result.get('sent_item'):
                result['index']=delivery.ensure_indexed(email,result.get('subject'),result.get('sent_at'))
        return result

    def refresh_indexes(self):
        from supervisor import schedule_index
        result=schedule_index(force=True)
        if self.imap.configs:self.imap_pool.submit(index_imap.refresh,self.imap,force=True)
        return result

    def unified_sync(self):
        self.imap_pool.submit(index_imap.refresh,self.imap,force=True)
        result=self.worker.submit('sync',{}).result(timeout=120)
        result['note']+=' Direct IMAP mailboxes read from the server; their search data is updated as well.'
        return result

    def job(self, request_id, action, args):
        fingerprint = json.dumps([action, args], sort_keys=True, ensure_ascii=False)
        with self.jobs_lock:
            if request_id in self.jobs:
                future, previous, _ = self.jobs[request_id]
                if previous != fingerprint:
                    raise MailError('This request code was already used for a different command.')
                return future
            # Retain the last 1,000 requests for duplicate prevention; pending ones never pruned.
            if len(self.jobs) >= 1000:
                completed = [k for k, v in self.jobs.items() if v[0].done()]
                for k in completed[:100]:
                    del self.jobs[k]
                if len(self.jobs) >= 1000:
                    raise MailError('Too many commands at once. Wait a moment.')
            if action=='quick_search':
                permitted={'query','email','limit','offset','include_deleted','sender','folder_id','after','before','unread'}
                if set(args)-permitted:raise MailError('Unknown search filters.')
                future=self.index_pool.submit(mail_index.combined_search,**args)
            elif action=='index_status':
                if args:raise MailError('index_status takes no arguments.')
                future=self.index_pool.submit(mail_index.combined_status)
            elif action=='refresh_index':
                if args:raise MailError('refresh_index takes no arguments.')
                future=self.index_pool.submit(self.refresh_indexes)
            elif action=='status' and self.imap.configs:
                if args:raise MailError('status takes no arguments.')
                future=self.router_pool.submit(self.unified_status)
            elif action=='verify_sent':
                permitted={'email','item_id','message_id','subject','expect_to','expect_attachments','deep'}
                if set(args)-permitted or not args.get('email'):raise MailError('verify_sent requires email and only accepts item_id, message_id, subject, expect_to, expect_attachments and deep.')
                future=self.router_pool.submit(self.verify_sent,args)
            elif action=='sync' and self.imap.configs:
                if args:raise MailError('sync takes no arguments.')
                future=self.router_pool.submit(self.unified_sync)
            elif action in ('update_draft','delete_draft') and not args.get('email'):
                raise MailError(action+' requires email.')
            elif str(args.get('email','')).lower() in self.imap.configs:
                if action not in {'folders','messages','read','draft','update_draft','delete_draft','send','search','suggest_folder','move','trash','signatures'}:
                    raise MailError('This command is not available for IMAP.')
                future=self.imap_pool.submit(getattr(self.imap,action),**args)
            elif action=='search' and not args.get('email') and self.imap.configs:
                future=self.router_pool.submit(self.full_search,args)
            else:
                future = self.worker.submit(action, args)
            self.jobs[request_id] = (future, fingerprint, time.monotonic())
            return future


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # Never log message content, tokens, or recipient addresses.

    def response(self, code, data, content_type='application/json; charset=utf-8'):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8') if isinstance(data, (dict, list)) else data
        self.send_response(code)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
        try:
            self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError,ConnectionResetError,ConnectionAbortedError):
            pass  # Closing the browser does not cancel or repeat its native job.

    def gate(self, authenticated=False):
        if self.headers.get('Host') != f'127.0.0.1:{self.server.server_port}':
            self.response(403, {'error': 'Only the local address is allowed.'})
            return False
        origin = self.headers.get('Origin')
        if origin and origin != self.server.origin:
            self.response(403, {'error': 'Another website has no access.'})
            return False
        if authenticated and not hmac.compare_digest(self.headers.get('X-Local-Token', ''), self.server.token):
            self.response(401, {'error': 'Invalid or missing local token.'})
            return False
        return True

    def do_OPTIONS(self):
        self.response(403, {'error': 'No access from other websites.'})

    def do_GET(self):
        if not self.gate():
            return
        path = urlsplit(self.path).path
        if path == '/health':
            self.response(200, {'service': 'outlook-local-manager', 'version': 2})
        else:
            self.response(404, {'error': 'Unknown route.'})

    def do_POST(self):
        if not self.gate(authenticated=True):
            return
        if urlsplit(self.path).path != '/api/call':
            self.response(404, {'error': 'Unknown route.'})
            return
        try:
            if self.headers.get('Content-Type', '').split(';')[0] != 'application/json':
                raise MailError('JSON required.')
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= 600000:
                raise MailError('Invalid request size.')
            payload = json.loads(self.rfile.read(size))
            action = payload.get('action')
            args = payload.get('args', {})
            request_id = payload.get('request_id', '')
            if not isinstance(args, dict):
                raise MailError('Invalid arguments.')
            if action == 'poll':
                with self.server.jobs_lock:
                    job = self.server.jobs.get(request_id)
                if not job:
                    raise MailError('The command is unknown or the service was restarted. Check Outlook before saving a draft again.')
                future = job[0]
            else:
                if action not in {'status', 'folders', 'messages', 'read', 'draft', 'update_draft', 'delete_draft', 'sync', 'signatures', 'send', 'search', 'suggest_folder', 'move', 'trash', 'quick_search', 'index_status', 'refresh_index', 'verify_sent'}:
                    raise MailError('Unknown command.')
                if not isinstance(request_id, str) or not 16 <= len(request_id) <= 100:
                    raise MailError('A unique request code is required.')
                future = self.server.job(request_id, action, args)
            try:
                self.response(200, future.result(timeout=10))
            except concurrent.futures.TimeoutError:
                self.response(202, {'pending': True, 'request_id': request_id})
        except (MailError, ValueError, TypeError) as exc:
            self.response(400, {'error': str(exc)})
        except Exception:
            self.response(500, {'error': 'The local command failed. Check Outlook and try reading again.'})


def main():
    RUNTIME.mkdir(exist_ok=True)
    server = Manager(('127.0.0.1', PORT))
    state = {'url': server.origin, 'token': server.token, 'pid': os.getpid(), 'version': 2}
    state_path = RUNTIME / 'session.json'
    state_path.write_text(json.dumps(state), encoding='utf-8')
    try:
        server.serve_forever()
    finally:
        server.server_close()
        state_path.unlink(missing_ok=True)


if __name__ == '__main__':
    main()
