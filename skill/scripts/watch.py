"""Read-only monitoring on top of the local mail service. Never sends, edits or deletes anything.

  watch.py drafts   [--stale-days 7]  Find duplicate and stale drafts across connected mailboxes.
  watch.py bounces  [--days 3]        Verify recent Sent items and report new bounces (NDRs).

Both print JSON and keep only metadata (hashes, subjects, dates) in bridge/runtime/watch-state.json.
Optional bridge/runtime/watch-config.json:
  {"ignore_drafts": [{"account": "you@example.com", "subject_regex": "Booking 12345"}]}
Matching drafts are never grouped, reported as stale, or suggested for deletion.
"""
import argparse
import datetime as dt
import difflib
import hashlib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mail  # noqa: E402  (sets up the bridge path, ensure_service and verify_sent)
from client import call  # noqa: E402

RUNTIME = mail.BRIDGE / 'runtime'
STATE = RUNTIME / 'watch-state.json'
CONFIG = RUNTIME / 'watch-config.json'
PREFIX = re.compile(r'^\s*((re|fw|fwd|aw|antw|wg|doorst)\s*:\s*)+', re.I)


def now():
    return dt.datetime.now(dt.timezone.utc)


def load(path, default):
    try:
        return json.loads(path.read_text('utf-8'))
    except (OSError, ValueError):
        return default


def save_state(state):
    RUNTIME.mkdir(exist_ok=True)
    temp = STATE.with_suffix('.tmp')
    temp.write_text(json.dumps(state, ensure_ascii=False), encoding='utf-8')
    temp.replace(STATE)


def age_days(value):
    try:
        return round((now() - dt.datetime.fromisoformat(value)).total_seconds() / 86400, 1)
    except (TypeError, ValueError):
        return None


def connected_accounts(only=None):
    status = call('status')
    accounts = [a['email'] for a in status['accounts'] if a['connected']]
    missing = [a['email'] for a in status['accounts'] if not a['connected']]
    if only:
        accounts = [a for a in accounts if a.lower() in {o.lower() for o in only}]
    return accounts, missing


def page(account, folder_id, cutoff=None, maximum=1000):
    """Messages of one folder, newest first, optionally stopping at a cutoff date."""
    rows, offset = [], 0
    while offset is not None and len(rows) < maximum:
        result = call('messages', {'email': account, 'folder_id': folder_id, 'limit': 200, 'offset': offset})
        for item in result['items']:
            if cutoff and item['received'] < cutoff:
                return rows
            rows.append(item)
        offset = result.get('next_offset')
    return rows


def ignored(config, account, subject):
    return any(r.get('account', '').lower() == account.lower() and re.search(r['subject_regex'], subject, re.I)
               for r in config.get('ignore_drafts', []))


def draft_audit(args):
    config = load(CONFIG, {})
    accounts, missing = connected_accounts(args.account)
    report = {'checked_at': now().isoformat(), 'missing_accounts': missing, 'duplicate_groups': [],
              'stale_drafts': [], 'ignored': 0, 'errors': [], 'sends_or_deletes_anything': False}
    for account in accounts:
        try:
            folder_id = call('folders', {'email': account})['drafts_id']
            drafts = page(account, folder_id)
        except Exception as exc:
            report['errors'].append({'account': account, 'error': str(exc)[:200]})
            continue
        groups = {}
        for draft in drafts:
            if ignored(config, account, draft['subject']):
                report['ignored'] += 1
                continue
            groups.setdefault(PREFIX.sub('', draft['subject']).strip().lower(), []).append(draft)
            days = age_days(draft['received'])
            if days is not None and days >= args.stale_days:
                report['stale_drafts'].append({'account': account, 'id': draft['id'], 'subject': draft['subject'],
                                               'age_days': days})
        for subject, members in groups.items():
            if len(members) < 2:
                continue
            details = []
            for draft in sorted(members, key=lambda d: d['received'], reverse=True):
                try:
                    full = call('read', {'email': account, 'item_id': draft['id']})
                    details.append({'id': draft['id'], 'subject': draft['subject'], 'received': draft['received'],
                                    'to': full.get('to', ''), 'attachments': [a['name'] for a in full.get('attachments', [])],
                                    '_body': (full.get('body') or '')[:5000]})
                except Exception as exc:
                    report['errors'].append({'account': account, 'id': draft['id'], 'error': str(exc)[:200]})
            if len(details) < 2:
                continue
            keep = details[0]
            group = {'account': account, 'subject': keep['subject'], 'keep_newest': keep['id'], 'drafts': []}
            for d in details:
                ratio = difflib.SequenceMatcher(None, keep['_body'], d['_body']).ratio()
                same_to = d['to'].lower() == keep['to'].lower()
                d['similarity_to_newest'] = round(ratio, 3)
                d['same_recipients'] = same_to
                d['likely_duplicate'] = d is not keep and same_to and ratio >= 0.9
                group['drafts'].append(d)
            for d in details:
                d.pop('_body')  # bodies never leave this function
            group['suggested_delete'] = [d['id'] for d in group['drafts'] if d['likely_duplicate']]
            group['note'] = 'Suggestion only. Delete a draft with delete-draft after the owner approves; different recipients or text means: review manually.'
            report['duplicate_groups'].append(group)
    report['summary'] = {'duplicate_groups': len(report['duplicate_groups']), 'stale_drafts': len(report['stale_drafts']),
                         'suggested_deletes': sum(len(g['suggested_delete']) for g in report['duplicate_groups'])}
    return report


def bounce_watch(args):
    state = load(STATE, {})
    seen = state.setdefault('sent', {})
    cutoff = (now() - dt.timedelta(days=args.days)).isoformat()
    accounts, missing = connected_accounts(args.account)
    report = {'checked_at': now().isoformat(), 'missing_accounts': missing, 'new_bounces': [], 'verified': 0,
              'pending': 0, 'delivered': 0, 'ambiguous_ndrs': [], 'skipped_over_limit': 0, 'errors': []}
    budget = args.max
    for account in accounts:
        try:
            folder_id = call('folders', {'email': account}).get('sent_id')
            if not folder_id:
                continue
            sent = page(account, folder_id, cutoff=cutoff, maximum=400)
        except Exception as exc:
            report['errors'].append({'account': account, 'error': str(exc)[:200]})
            continue
        keyed = [(hashlib.sha256((account.lower() + '|' + item['id']).encode()).hexdigest()[:24], item) for item in sent]
        keyed.sort(key=lambda row: row[0] in seen)  # never-checked messages first, so the cap cannot starve them
        for key, item in keyed:
            entry = seen.get(key)
            if entry and entry['status'] in ('delivered', 'bounced', 'settled'):
                continue
            if entry and entry.get('last_check', '') > (now() - dt.timedelta(hours=args.recheck_hours)).isoformat():
                report['pending'] += 1
                continue
            if budget <= 0:
                report['skipped_over_limit'] += 1
                continue
            budget -= 1
            try:
                result = mail.verify_sent({'email': account, 'item_id': item['id']}, 0)
            except Exception as exc:
                report['errors'].append({'account': account, 'subject': item['subject'][:80], 'error': str(exc)[:200]})
                continue
            report['verified'] += 1
            status = 'bounced' if result.get('bounce_detected') else 'delivered' if result.get('delivery_verified') else 'pending'
            for ndr in result.get('possible_bounces', []):
                # An NDR that cannot be tied to one message (same subject/time, e.g. in a mail merge): report it once.
                ndr_key = hashlib.sha256(f"{account}|{ndr.get('received')}|{ndr.get('subject')}".encode()).hexdigest()[:24]
                if ndr_key not in state.setdefault('ndrs', {}):
                    state['ndrs'][ndr_key] = {'seen': now().isoformat()}
                    report['ambiguous_ndrs'].append({'account': account, 'ndr_subject': ndr.get('subject'),
                                                     'ndr_received': ndr.get('received'), 'folder': ndr.get('folder_path'),
                                                     'note': 'Cannot be tied to a single sent message; open the NDR to see which recipient failed.'})
            entry = seen.setdefault(key, {'subject': item['subject'][:120], 'first_seen': now().isoformat(), 'alerted': False})
            entry['status'] = status
            entry['last_check'] = now().isoformat()
            if status == 'bounced' and not entry['alerted']:
                entry['alerted'] = True
                report['new_bounces'].append({'account': account, 'subject': item['subject'], 'sent': item['received'],
                                              'failed_recipients': result.get('failed_recipients', []),
                                              'explanation': result.get('explanation')})
            elif status == 'delivered':
                report['delivered'] += 1
            elif status == 'pending':
                report['pending'] += 1
    # Anything older than the window can no longer be re-checked here: settle it so state stays small.
    horizon = (now() - dt.timedelta(days=args.days * 3)).isoformat()
    state['sent'] = {k: v for k, v in seen.items() if v.get('first_seen', '') >= horizon}
    state['ndrs'] = {k: v for k, v in state.get('ndrs', {}).items() if v.get('seen', '') >= horizon}
    save_state(state)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest='command', required=True)
    drafts = commands.add_parser('drafts')
    drafts.add_argument('--stale-days', type=float, default=7)
    bounces = commands.add_parser('bounces')
    bounces.add_argument('--days', type=float, default=3, help='Look at Sent items from the last N days')
    bounces.add_argument('--recheck-hours', type=float, default=3, help='Do not re-verify a pending message more often than this')
    bounces.add_argument('--max', type=int, default=40, help='Maximum verify-sent checks per run')
    for sub in (drafts, bounces):
        sub.add_argument('--account', action='append', help='Limit to this mailbox (repeatable)')
    args = parser.parse_args()
    try:
        mail.ensure_service()
        result = draft_audit(args) if args.command == 'drafts' else bounce_watch(args)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
