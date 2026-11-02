"""Assistant CLI for the local Outlook/IMAP mail service. Windows only; prints JSON.

Run it with the Python environment that has the bridge requirements installed
(pywin32), for example:  .venv\\Scripts\\python.exe skill\\scripts\\mail.py status
Set OUTLOOK_LOCAL_BRIDGE to the bridge directory if the repository layout differs.
"""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from urllib.request import urlopen

if sys.platform != 'win32':
    sys.exit('outlook-local-mail only works on Windows (Outlook Classic COM + local loopback service).')

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

BRIDGE = Path(os.environ.get('OUTLOOK_LOCAL_BRIDGE') or Path(__file__).resolve().parents[2] / 'bridge')
sys.path.insert(0, str(BRIDGE))
from client import call


def ensure_service():
    try:
        with urlopen('http://127.0.0.1:38473/health', timeout=2) as response:
            if json.load(response).get('service') != 'outlook-local-manager':
                raise RuntimeError('Port 38473 is used by another program.')
        return
    except (ConnectionError, OSError):
        # DEVNULL lets us wait only for the starter, not for the service lifetime.
        pythonw = Path(sys.executable).with_name('pythonw.exe')
        result = subprocess.run([str(pythonw), str(BRIDGE / 'supervisor.py')], stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode:
            raise RuntimeError('The local mail service did not start. Check bridge/runtime/service-error.log.')


def overview(limit):
    state = call('status')
    result = {'checked_at': state['checked_at'], 'connected': state['connected'],
              'expected': state['expected'], 'offline': state['offline'],
              'sync_verified': state['sync_verified'], 'error': state['error'],
              'missing': [a['email'] for a in state['accounts'] if not a['connected']], 'mailboxes': []}
    for account in state['accounts']:
        if account['connected']:
            row = {'email': account['email']}
            try:
                row.update(call('messages', {'email': account['email'], 'limit': limit}))
            except Exception as exc:
                row['error'] = str(exc)
            result['mailboxes'].append(row)
    return result


def verify_sent(args, wait=0, interval=20):
    """Poll verify_sent outside the quick index until delivered/bounced or the wait ends."""
    started = time.monotonic()
    while True:
        result = call('verify_sent', args, polls=400)
        waited = round(time.monotonic() - started)
        settled = result.get('bounce_detected') or result.get('delivery_verified')
        if settled or waited + interval > wait:
            result['waited_seconds'] = waited
            return result
        time.sleep(interval)


def main():
    parser = argparse.ArgumentParser(description='Local Outlook Classic + IMAP mail for AI agents')
    actions = parser.add_subparsers(dest='action', required=True)
    for name in ('status', 'sync', 'index-status', 'refresh-index'):
        actions.add_parser(name)
    quick = actions.add_parser('quick-search')
    quick.add_argument('--query', default='')
    quick.add_argument('--account')
    quick.add_argument('--sender', default='')
    quick.add_argument('--folder')
    quick.add_argument('--after')
    quick.add_argument('--before')
    quick.add_argument('--unread', action='store_true', default=None)
    quick.add_argument('--include-deleted', action='store_true')
    quick.add_argument('--limit', type=int, default=50)
    quick.add_argument('--offset', type=int, default=0)
    for action in ('folders', 'signatures'):
        folders = actions.add_parser(action)
        folders.add_argument('--account', required=True)
        if action == 'folders':
            folders.add_argument('--refresh', action='store_true')
    search = actions.add_parser('search')
    search.add_argument('--query', required=True)
    search.add_argument('--account')
    search.add_argument('--limit', type=int, default=50)
    search.add_argument('--offset', type=int, default=0)
    search.add_argument('--include-deleted', action='store_true')
    search.add_argument('--cursor')
    advice = actions.add_parser('suggest-folder')
    advice.add_argument('--account', required=True)
    advice.add_argument('--id', required=True)
    trash = actions.add_parser('trash')
    trash.add_argument('--account', required=True)
    trash.add_argument('--id', required=True)
    trash.add_argument('--source-folder', required=True)
    trash.add_argument('--reason', required=True)
    move = actions.add_parser('move')
    move.add_argument('--account', required=True)
    move.add_argument('--id', required=True)
    move.add_argument('--folder', required=True)
    move.add_argument('--source-folder')
    messages = actions.add_parser('messages')
    messages.add_argument('--account', required=True)
    messages.add_argument('--folder')
    messages.add_argument('--query', default='')
    messages.add_argument('--offset', type=int, default=0)
    messages.add_argument('--limit', type=int, default=50)
    read = actions.add_parser('read')
    read.add_argument('--account', required=True)
    read.add_argument('--id', required=True)
    draft = actions.add_parser('draft')
    draft.add_argument('--file', required=True)
    update = actions.add_parser('update-draft')
    update.add_argument('--account', required=True)
    update.add_argument('--id', required=True, help='ID of the existing draft in Drafts')
    update.add_argument('--file', required=True, help='JSON with only the fields to change: subject, body, to, cc, bcc, attachments_add, attachments_remove')
    remove = actions.add_parser('delete-draft')
    remove.add_argument('--account', required=True)
    remove.add_argument('--id', required=True)
    remove.add_argument('--reason', required=True)
    remove.add_argument('--expect-subject', help='Refuse unless the draft subject matches this exactly')
    send = actions.add_parser('send')
    send.add_argument('--account', required=True)
    send.add_argument('--id', required=True)
    send.add_argument('--verify-wait', type=int, default=None,
                      help='After sending, run verify-sent and wait up to this many seconds for a bounce/receipt')
    verify = actions.add_parser('verify-sent')
    verify.add_argument('--account', required=True)
    verify.add_argument('--id', help='Draft ID given to send, or the ID of the sent message')
    verify.add_argument('--message-id')
    verify.add_argument('--subject')
    verify.add_argument('--expect-to', help='Expected recipients, comma-separated')
    verify.add_argument('--expect-attachments', type=int, help='Expected number of ordinary (non-inline) attachments')
    verify.add_argument('--deep', action='store_true', help='Look for bounces in the whole mailbox, not just Inbox/Junk/Deleted')
    verify.add_argument('--wait', type=int, default=0, help='Repeat for at most this many seconds (0..900)')
    verify.add_argument('--interval', type=int, default=20)
    summary = actions.add_parser('overview')
    summary.add_argument('--limit', type=int, default=15)
    args = parser.parse_args()
    try:
        if hasattr(args, 'limit') and not 1 <= args.limit <= 200:
            parser.error('--limit must be between 1 and 200')
        ensure_service()
        if args.action == 'overview':
            result = overview(args.limit)
        elif args.action == 'index-status':
            result = call('index_status')
        elif args.action == 'refresh-index':
            result = call('refresh_index')
        elif args.action == 'quick-search':
            result = call('quick_search', {'query': args.query, 'email': args.account, 'sender': args.sender,
                'folder_id': args.folder, 'after': args.after, 'before': args.before, 'unread': args.unread,
                'limit': args.limit, 'offset': args.offset, 'include_deleted': args.include_deleted})
        elif args.action in ('folders', 'signatures'):
            params = {'email': args.account}
            if args.action == 'folders':
                params['refresh'] = args.refresh
            result = call(args.action, params)
        elif args.action == 'search':
            params = {'email': args.account, 'query': args.query, 'limit': args.limit, 'offset': args.offset,
                      'include_deleted': args.include_deleted, 'cursor': args.cursor}
            result = call('search', params)
            # The service resumes a bounded scan with the same snapshot; don't
            # present early folder batches as complete mailbox search results.
            for _ in range(19):
                if result['scan_complete']:
                    break
                params['cursor'] = result['cursor']
                result = call('search', params)
        elif args.action == 'suggest-folder':
            result = call('suggest_folder', {'email': args.account, 'item_id': args.id})
        elif args.action == 'trash':
            result = call('trash', {'email': args.account, 'item_id': args.id,
                                    'source_folder_id': args.source_folder, 'reason': args.reason})
        elif args.action == 'move':
            result = call('move', {'email': args.account, 'item_id': args.id, 'folder_id': args.folder,
                                   'source_folder_id': args.source_folder})
        elif args.action == 'messages':
            result = call('messages', {'email': args.account, 'folder_id': args.folder, 'query': args.query,
                                       'offset': args.offset, 'limit': args.limit})
        elif args.action == 'read':
            result = call('read', {'email': args.account, 'item_id': args.id})
        elif args.action == 'draft':
            result = call('draft', json.loads(Path(args.file).read_text('utf-8-sig')))
        elif args.action == 'update-draft':
            fields = json.loads(Path(args.file).read_text('utf-8-sig'))
            if not isinstance(fields, dict) or {'email', 'item_id', 'id'} & set(fields):
                parser.error('--file contains only the fields to change; pass the account and draft ID with --account and --id')
            result = call('update_draft', dict(fields, email=args.account, item_id=args.id))
        elif args.action == 'delete-draft':
            params = {'email': args.account, 'item_id': args.id, 'reason': args.reason}
            if args.expect_subject is not None:
                params['expect_subject'] = args.expect_subject
            result = call('delete_draft', params)
        elif args.action == 'send':
            result = call('send', {'email': args.account, 'item_id': args.id})
            if args.verify_wait is not None and result.get('submitted'):
                if not 0 <= args.verify_wait <= 900:
                    parser.error('--verify-wait must be between 0 and 900')
                time.sleep(min(15, args.verify_wait))
                result = {'send': result, 'verify': verify_sent({'email': args.account, 'item_id': args.id},
                                                                max(0, args.verify_wait - 15))}
        elif args.action == 'verify-sent':
            if not (args.id or args.message_id or args.subject):
                parser.error('Give --id, --message-id or --subject')
            if not 0 <= args.wait <= 900 or not 5 <= args.interval <= 300:
                parser.error('--wait 0..900 and --interval 5..300')
            params = {'email': args.account, 'deep': args.deep}
            for key, value in (('item_id', args.id), ('message_id', args.message_id), ('subject', args.subject),
                               ('expect_to', args.expect_to), ('expect_attachments', args.expect_attachments)):
                if value is not None:
                    params[key] = value
            result = verify_sent(params, args.wait, args.interval)
        else:
            result = call(args.action)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
