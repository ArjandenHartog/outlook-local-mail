"""Post-send verification: Sent/Outbox state, bounces, receipts and own-mailbox arrival.

Pure helpers are separate from the Outlook adapter so they can be unit tested.
No bodies are stored; report bodies are only scanned in memory for the Message-ID.
"""
import datetime as dt
import hashlib
import json
import re
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
P = 'http://schemas.microsoft.com/mapi/proptag/'
MSGID = P + '0x1035001F'
SUBMIT = P + '0x00390040'
SEARCH_KEY = P + '0x300B0102'
SIZE = P + '0x0E080003'
RECEIVED = P + '0x0E060040'
CLASS = P + '0x001A001F'
SENDER_SMTP = P + '0x5D01001F'
ORIGINAL_SUBJECT = P + '0x0049001F'
ORIGINAL_SUBMIT = P + '0x004E0040'
RECIPIENT_SMTP = P + '0x39FE001F'
ATTACH_CONTENT_ID = P + '0x3712001F'
ATTACH_HIDDEN = P + '0x7FFE000B'
DR_REQUESTED = P + '0x0023000B'
RR_REQUESTED = P + '0x0029000B'

# Candidate filter for report-like mail. The final decision is made in Python.
REPORT_SUBJECTS = ('Undeliverable', 'Onbestelbaar', 'Niet bezorgd', 'Niet afgeleverd', 'Delivery Status Notification',
                   'Mail delivery failed', 'Returned mail', 'Undelivered Mail', 'Delivery Failure', 'failure notice',
                   'Delivered:', 'Bezorgd:', 'Afgeleverd:', 'Read:', 'Gelezen:', 'Not read:', 'Niet gelezen:',
                   'Delivery delayed', 'Bezorging vertraagd', 'Vertraagd:')
REPORT_SENDERS = ('postmaster', 'mailer-daemon', 'microsoftexchange')
FAILURE_SUBJECT = re.compile(r'(?i)\b(undeliverable|onbestelbaar|niet bezorgd|niet afgeleverd|delivery status notification \(failure\)|'
                             r'mail delivery failed|returned mail|undelivered mail|delivery failure|failure notice)')
DELAY_SUBJECT = re.compile(r'(?i)(delivery delayed|bezorging vertraagd|vertraagd:|delayed)')
DELIVERED_SUBJECT = re.compile(r'(?i)^(delivered|bezorgd|afgeleverd)\s*:')
READ_SUBJECT = re.compile(r'(?i)^(read|gelezen)\s*:')
NOT_READ_SUBJECT = re.compile(r'(?i)^(not read|niet gelezen)\s*:')
MESSAGE_ID = re.compile(r'<[^<>\s]+@[^<>\s]+>')


def utcnow():
    return dt.datetime.now(dt.timezone.utc)


def iso(value):
    if value is None:
        return None
    if isinstance(value, str):
        return value
    try:
        return value.astimezone(dt.timezone.utc).isoformat()
    except (ValueError, AttributeError, OSError):
        return str(value)


def parse_time(value):
    if not value:
        return None
    if isinstance(value, dt.datetime):
        return value if value.tzinfo else value.replace(tzinfo=dt.timezone.utc)
    try:
        parsed = dt.datetime.fromisoformat(str(value))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=dt.timezone.utc)
    except ValueError:
        return None


def hexkey(value):
    if value is None:
        return None
    if isinstance(value, (bytes, bytearray, memoryview)):
        return bytes(value).hex().lower()
    return str(value).replace(' ', '').lower() or None


def send_key(email, item_id):
    return hashlib.sha256((str(email).lower() + '\0' + str(item_id)).encode()).hexdigest()


def normalize_addresses(values):
    if values is None:
        return []
    if isinstance(values, str):
        values = re.split(r'[;,]', values)
    return sorted({str(v).strip().lower() for v in values if str(v).strip()})


def classify_report(message_class='', subject='', sender='', body=''):
    """Return ndr/delivery_receipt/read_receipt/not_read/delayed or None."""
    cls = str(message_class or '').upper()
    subject = str(subject or '')
    sender = str(sender or '').lower()
    if cls.startswith('REPORT.'):
        if cls.endswith('.NDR'):
            return 'ndr'
        if cls.endswith('.DR'):
            return 'delivery_receipt'
        if cls.endswith('.IPNRN'):
            return 'read_receipt'
        if cls.endswith('.IPNNRN'):
            return 'not_read'
        if 'DELAY' in cls:
            return 'delayed'
    lowered = str(body or '')[:20000].lower()
    if 'action: failed' in lowered or FAILURE_SUBJECT.search(subject):
        return 'ndr'
    if 'action: delayed' in lowered or DELAY_SUBJECT.search(subject) and any(s in sender for s in REPORT_SENDERS):
        return 'delayed'
    if 'action: delivered' in lowered or 'action: relayed' in lowered or DELIVERED_SUBJECT.search(subject):
        return 'delivery_receipt'
    if 'disposition:' in lowered and 'displayed' in lowered or READ_SUBJECT.search(subject):
        return 'read_receipt'
    if NOT_READ_SUBJECT.search(subject):
        return 'not_read'
    if any(s in sender for s in ('postmaster', 'mailer-daemon')):
        return 'ndr'
    return None


def strip_prefix(subject):
    return re.sub(r'(?i)^\s*((undeliverable|onbestelbaar|niet bezorgd|delivered|bezorgd|afgeleverd|read|gelezen|'
                  r'not read|niet gelezen|delivery delayed|vertraagd)\s*:\s*)+', '', str(subject or '')).strip()


def match_report(report, target):
    """How strongly does a report refer to the target message?"""
    body = str(report.get('body') or '')
    message_id = str(target.get('message_id') or '')
    if message_id and message_id.lower() in body.lower():
        return 'message_id'
    subject = str(target.get('subject') or '').strip()
    original_subject = str(report.get('original_subject') or '').strip()
    original_submit = parse_time(report.get('original_submit'))
    sent_at = parse_time(target.get('sent_at'))
    if subject and original_subject == subject and original_submit and sent_at \
            and abs((original_submit - sent_at).total_seconds()) <= 120:
        return 'original_subject_time'
    if subject and (strip_prefix(report.get('subject')) == subject or original_subject == subject):
        return 'subject_only'
    return None


def recipients_in(body, recipients):
    lowered = str(body or '').lower()
    return [r for r in recipients if r and r.lower() in lowered]


def report_row(report, kind, match, recipients):
    return {'kind': kind, 'match': match, 'id': report.get('id'), 'folder_path': report.get('folder_path'),
            'subject': report.get('subject'), 'sender': report.get('sender'), 'received': iso(report.get('received')),
            'recipients': recipients_in(report.get('body'), recipients)}


def evaluate_reports(reports, target, recipients):
    """Split report candidates into strong matches and weak (subject-only) hints."""
    found = {'ndr': [], 'delivery_receipt': [], 'read_receipt': [], 'not_read': [], 'delayed': []}
    weak = []
    for report in reports:
        kind = classify_report(report.get('message_class'), report.get('subject'), report.get('sender'), report.get('body'))
        if not kind:
            continue
        match = match_report(report, target)
        if not match:
            continue
        row = report_row(report, kind, match, recipients)
        (weak if match == 'subject_only' else found[kind]).append(row)
    return found, weak


def compare_expected(expected, actual):
    """Compare intended subject/recipients/attachments with the Sent copy."""
    checks = {}
    if expected.get('subject') is not None:
        checks['subject'] = str(expected['subject']).strip() == str(actual.get('subject') or '').strip()
    if expected.get('recipients'):
        want = set(normalize_addresses(expected['recipients']))
        have = set(normalize_addresses(r['address'] for r in actual.get('recipients', [])))
        checks['recipients'] = want == have
        if want != have:
            checks['recipients_missing'] = sorted(want - have)
            checks['recipients_unexpected'] = sorted(have - want)
    regular = [a for a in actual.get('attachments', []) if not a.get('inline')]
    if expected.get('attachments') is not None:
        want = sorted(str(a['name']) for a in expected['attachments'])
        have = sorted(str(a['name']) for a in actual.get('attachments', []))
        checks['attachments'] = want == have
        if want != have:
            checks['attachments_missing'] = sorted(set(want) - set(have))
            checks['attachments_unexpected'] = sorted(set(have) - set(want))
        sizes = {a['name']: a.get('size') for a in expected['attachments']}
        differences = [{'name': a['name'], 'draft_size': sizes[a['name']], 'sent_size': a.get('size')}
                       for a in actual.get('attachments', []) if a['name'] in sizes and sizes[a['name']] != a.get('size')]
        if differences:
            # Outlook's attachment Size includes MAPI overhead; small drift is normal.
            checks['attachment_size_differences'] = differences
            checks['attachment_sizes_within_5_percent'] = all(
                d['draft_size'] and d['sent_size'] and abs(d['sent_size'] - d['draft_size']) <= 0.05 * d['draft_size']
                for d in differences)
    if expected.get('attachment_count') is not None:
        checks['attachment_count'] = int(expected['attachment_count']) == len(regular)
    return checks


def finalize(result):
    """Derive delivery status fields and a Dutch explanation. Idempotent."""
    recipients = [r['address'] for r in result.get('recipients', []) if r.get('address')]
    found = result.get('reports', {})
    bounces = found.get('ndr', [])
    result['bounce_detected'] = bool(bounces)
    result['bounces'] = bounces
    result['possible_bounces'] = [r for r in result.get('weak_reports', []) if r['kind'] == 'ndr']
    result['delayed'] = found.get('delayed', [])
    receipts = found.get('delivery_receipt', [])
    reads = found.get('read_receipt', [])
    result['delivery_receipt'] = {'requested': bool(result.get('delivery_receipt_requested')), 'received': receipts}
    result['read_receipt'] = {'requested': bool(result.get('read_receipt_requested')), 'received': reads,
                              'not_read': found.get('not_read', [])}
    delivered = {}
    for row in result.get('recipient_mailboxes', []):
        if row.get('found'):
            delivered.setdefault(row['recipient'], []).append('own_mailbox')
    for source, rows in (('delivery_receipt', receipts), ('read_receipt', reads)):
        for row in rows:
            named = row['recipients'] or (recipients if len(recipients) == 1 else [])
            for address in named:
                if source not in delivered.setdefault(address, []):
                    delivered[address].append(source)
    failed = {a for row in bounces for a in (row['recipients'] or recipients)}
    for address in failed:
        delivered.pop(address, None)
    result['delivered_recipients'] = [{'recipient': a, 'evidence': delivered[a]} for a in recipients if a in delivered]
    result['unverified_recipients'] = [a for a in recipients if a not in delivered and a not in failed]
    result['failed_recipients'] = sorted(failed)
    result['delivery_verified'] = bool(recipients) and not bounces and all(a in delivered for a in recipients)
    sent_at = parse_time(result.get('sent_at'))
    if sent_at:
        result['seconds_since_send'] = round((utcnow() - sent_at).total_seconds())
    if not result.get('in_sent_items'):
        state = 'in_outbox' if result.get('outbox_empty') is False else 'not_found'
    elif bounces:
        state = 'bounced'
    elif result['delivery_verified']:
        state = 'delivered'
    elif result['delivered_recipients']:
        state = 'partially_verified'
    else:
        state = 'sent_no_bounce_seen'
    result['state'] = state
    parts = []
    if state == 'in_outbox':
        parts.append('The message is still in the Outbox, so the server has not accepted it yet.')
    elif state == 'not_found':
        parts.append('No copy found in Sent Items for this submission; check again later or search by Message-ID/subject.')
    else:
        parts.append('The message is in Sent Items' + (' and the Outbox has no entry for it.' if result.get('outbox_empty') else '.'))
    if bounces:
        parts.append('A non-delivery report (NDR) referring to this message was found: delivery failed for '
                     + (', '.join(result['failed_recipients']) or 'at least one recipient') + '.')
    elif result.get('in_sent_items'):
        minutes = round(result.get('seconds_since_send', 0) / 60)
        scan = result.get('bounce_scan', {})
        parts.append(f"No bounce seen ({minutes} min after sending; {scan.get('folders_scanned', 0)} folders checked: "
                     f"{scan.get('scope', '-')}).")
        if not result['delivery_verified']:
            parts.append('The absence of a bounce is not proof of delivery; an NDR may still arrive later.')
    if result['possible_bounces']:
        parts.append('An NDR with the same subject but no Message-ID reference exists; review it manually.')
    labels = {'own_mailbox': 'the message is in that own mailbox (same Message-ID)',
              'delivery_receipt': 'delivery receipt from the receiving server', 'read_receipt': 'read receipt'}
    for row in result['delivered_recipients']:
        parts.append(f"Delivery to {row['recipient']} demonstrated: " +'; '.join(labels[e] for e in row['evidence']) + '.')
    if result['unverified_recipients'] and result.get('in_sent_items') and not bounces:
        requested = result['delivery_receipt']['requested'] or result['read_receipt']['requested']
        parts.append('Delivery not demonstrated to: ' + ', '.join(result['unverified_recipients']) + '. '
                     + ('A receipt was requested but has not (yet) been received.' if requested else
                        'No delivery or read receipt was requested and the recipient is not an own mailbox.'))
    result['explanation'] = ' '.join(parts)
    return result


def ensure_indexed(email, subject, sent_at, db_path=None, schedule=None, invalidate=None):
    """If the metadata index misses this sent item, mark it stale and request a rescan.

    Outlook tables return short-term EntryIDs, items long-term ones, so match on
    mailbox + Sent kind + subject + time instead of the ID.
    """
    import mail_index
    path = Path(db_path or mail_index.DB)
    contains = False
    moment = parse_time(sent_at)
    try:
        connection = mail_index.read_connection(path)
        try:
            rows = connection.execute("SELECT received FROM mail WHERE email=? AND kind='sent' AND subject=?",
                                      (str(email).lower(), str(subject or ''))).fetchall()
        finally:
            connection.close()
        contains = any(moment and parse_time(r[0]) and abs((parse_time(r[0]) - moment).total_seconds()) <= 600 for r in rows)
    except Exception:
        contains = False
    refresh = False
    if not contains:
        try:
            (invalidate or (lambda: mail_index.invalidate('verify-sent')))()
            if schedule is None:
                from supervisor import schedule_index
                schedule = schedule_index
            schedule()
            refresh = True
        except Exception:
            refresh = False
    return {'contains_sent_item': bool(contains), 'refresh_requested': refresh,
            'note': None if contains else 'The fast index does not contain this sent message yet; a silent re-index was requested. verify-sent works independently of the index.'}


class OutlookDeliveryMixin:
    """Mixed into OutlookBackend; all calls run on the STA worker thread."""

    @staticmethod
    def _prop(item, name, default=None):
        try:
            return item.PropertyAccessor.GetProperty(name)
        except Exception:
            return default

    def _recipient_address(self, recipient):
        address = self._prop(recipient, RECIPIENT_SMTP)
        if isinstance(address, str) and '@' in address:
            return address.lower()
        try:
            entry = recipient.AddressEntry
            if str(entry.Type) == 'EX':
                user = entry.GetExchangeUser()
                if user is not None and user.PrimarySmtpAddress:
                    return str(user.PrimarySmtpAddress).lower()
        except Exception:
            pass
        return str(getattr(recipient, 'Address', '') or '').lower()

    def _recipients(self, item):
        rows = []
        try:
            for recipient in item.Recipients:
                rows.append({'address': self._recipient_address(recipient), 'name': str(recipient.Name or ''),
                             'type': {1: 'to', 2: 'cc', 3: 'bcc'}.get(int(recipient.Type), str(recipient.Type))})
        except Exception:
            pass
        return rows

    def _attachments(self, item):
        rows = []
        try:
            for attachment in item.Attachments:
                content_id = self._prop(attachment, ATTACH_CONTENT_ID)
                hidden = bool(self._prop(attachment, ATTACH_HIDDEN, False))
                rows.append({'name': str(attachment.FileName), 'size': int(attachment.Size),
                             'inline': bool(hidden or content_id)})
        except Exception:
            pass
        return rows

    def fingerprint(self, item):
        """What we intended to send; recorded in the send journal before Send()."""
        result = {'subject': str(getattr(item, 'Subject', '') or ''), 'recipients': [], 'attachments': [],
                  'search_key': None, 'delivery_receipt_requested': False, 'read_receipt_requested': False}
        try:
            result['recipients'] = self._recipients(item)
            result['attachments'] = self._attachments(item)
            result['search_key'] = hexkey(self._prop(item, SEARCH_KEY))
            result['delivery_receipt_requested'] = bool(getattr(item, 'OriginatorDeliveryReportRequested', False))
            result['read_receipt_requested'] = bool(getattr(item, 'ReadReceiptRequested', False))
            message_id = self._prop(item, MSGID)
            if isinstance(message_id, str) and message_id:
                result['message_id'] = message_id
        except Exception:
            pass
        return result

    def _table(self, folder, restriction, columns):
        table = folder.GetTable(restriction)
        table.Columns.RemoveAll()
        for column in columns:
            table.Columns.Add(column)
        count = table.GetRowCount()
        return [list(row) for row in table.GetArray(min(count, 500))] if count else []

    def _sent_details(self, item):
        sent_at = self._prop(item, SUBMIT)
        return {'id': str(item.EntryID), 'folder_path': str(item.Parent.FolderPath),
                'subject': str(item.Subject or ''), 'message_id': self._prop(item, MSGID),
                'sent_at': iso(sent_at) if sent_at is not None else None,
                'size': self._prop(item, SIZE), 'search_key': hexkey(self._prop(item, SEARCH_KEY)),
                'recipients': self._recipients(item), 'attachments': self._attachments(item),
                'delivery_receipt_requested': bool(self._prop(item, DR_REQUESTED, False)),
                'read_receipt_requested': bool(self._prop(item, RR_REQUESTED, False))}

    def _find_sent(self, session, store_id, sent, message_id=None, subject=None, search_key=None, after=None):
        """Locate the Sent copy. Message-ID > search key > subject (+time)."""
        columns = ('EntryID', MSGID, SEARCH_KEY, SUBMIT)
        if message_id:
            literal = str(message_id).replace("'", "''")
            rows = self._table(sent, f'@SQL="{MSGID}" = \'{literal}\'', columns)
            return [r[0] for r in rows], 'message_id'
        if not subject:
            return [], None
        literal = str(subject).replace("'", "''")
        rows = self._table(sent, f'@SQL="urn:schemas:httpmail:subject" = \'{literal}\'', columns)
        if search_key:
            exact = [r for r in rows if hexkey(r[2]) == search_key]
            if exact:
                return [r[0] for r in exact], 'search_key'
        boundary = parse_time(after)
        if boundary:
            rows = [r for r in rows if parse_time(iso(r[3])) and parse_time(iso(r[3])) >= boundary - dt.timedelta(minutes=2)]
        rows.sort(key=lambda r: iso(r[3]) or '', reverse=True)
        return [r[0] for r in rows], 'subject_time' if boundary else 'subject'

    def _outbox(self, binding, subject=None, search_key=None, message_id=None):
        outbox = self.default_folder(binding, 4)
        matching = []
        count = 0
        try:
            for item in outbox.Items:
                count += 1
                key = hexkey(self._prop(item, SEARCH_KEY))
                same = (search_key and key == search_key) or (message_id and self._prop(item, MSGID) == message_id) \
                    or (not search_key and subject and str(item.Subject or '') == subject)
                if same:
                    matching.append({'id': str(item.EntryID), 'subject': str(item.Subject or '')})
        except Exception:
            pass
        return {'count': count, 'matching': matching}

    def _walk(self, folder, depth=0, limit=400):
        stack = [(folder, depth)]
        seen = 0
        while stack and seen < limit:
            current, level = stack.pop()
            seen += 1
            yield current
            if level < 25:
                try:
                    stack.extend((child, level + 1) for child in current.Folders)
                except Exception:
                    pass

    def _report_candidates(self, session, binding, since, deep=False):
        """Report-like items received since `since` in Inbox (+subfolders), Junk, Deleted, or the whole store."""
        store = binding[0]
        roots = [binding[0].GetRootFolder()] if deep else [self.default_folder(binding, n) for n in (6, 23, 3)]
        subjects = ' OR '.join(f'"urn:schemas:httpmail:subject" like \'%{s}%\'' for s in REPORT_SUBJECTS)
        senders = ' OR '.join(f'"{SENDER_SMTP}" like \'%{s}%\' OR "urn:schemas:httpmail:fromemail" like \'%{s}%\'' for s in REPORT_SENDERS)
        restriction = f'@SQL=("{CLASS}" like \'REPORT.%\') OR {subjects} OR {senders}'
        boundary = parse_time(since)
        candidates, folders, errors = [], [], []
        deadline = time.monotonic() + 25
        visited = set()
        for root in roots:
            for folder in self._walk(root):
                if time.monotonic() > deadline:
                    errors.append({'error': 'Bounce check time limit reached; not all folders were inspected.'})
                    break
                key = str(folder.EntryID)
                if key in visited:
                    continue
                visited.add(key)
                try:
                    if folder.DefaultItemType != 0:
                        continue
                    folders.append(str(folder.FolderPath))
                    for row in self._table(folder, restriction, ('EntryID', 'MessageClass', 'Subject', RECEIVED)):
                        received = parse_time(iso(row[3]))
                        if boundary and received and received < boundary - dt.timedelta(minutes=1):
                            continue
                        item = session.GetItemFromID(row[0], store.StoreID)
                        candidates.append({'id': str(row[0]), 'message_class': str(row[1] or ''), 'subject': str(row[2] or ''),
                                           'received': iso(row[3]), 'folder_path': str(folder.FolderPath),
                                           'sender': str(self._prop(item, SENDER_SMTP) or getattr(item, 'SenderEmailAddress', '') or ''),
                                           'body': str(getattr(item, 'Body', '') or '')[:200000],
                                           'original_subject': self._prop(item, ORIGINAL_SUBJECT),
                                           'original_submit': iso(self._prop(item, ORIGINAL_SUBMIT))})
                except Exception as exc:
                    errors.append({'folder': str(getattr(folder, 'FolderPath', '?')), 'error': str(exc)[:200]})
        return candidates, folders, errors

    def locate_message(self, email, message_id):
        """Find a Message-ID anywhere in an own Outlook mailbox (proof of arrival there)."""
        session, binding = self.mailbox(email)
        literal = str(message_id).replace("'", "''")
        restriction = f'@SQL="{MSGID}" = \'{literal}\''
        sent_id = str(self.default_folder(binding, 5).EntryID)
        found, errors, scanned = [], [], 0
        deadline = time.monotonic() + 25
        for folder in self._walk(binding[0].GetRootFolder()):
            if time.monotonic() > deadline:
                errors.append({'error': 'Tijdgrens bereikt.'})
                break
            try:
                if folder.DefaultItemType != 0 or str(folder.EntryID) == sent_id:
                    continue
                scanned += 1
                for row in self._table(folder, restriction, ('EntryID', RECEIVED)):
                    found.append({'id': str(row[0]), 'folder_path': str(folder.FolderPath), 'received': iso(row[1])})
            except Exception as exc:
                errors.append({'folder': str(getattr(folder, 'FolderPath', '?')), 'error': str(exc)[:200]})
        return {'email': str(email).lower(), 'message_id': message_id, 'found': bool(found), 'items': found,
                'scanned_folders': scanned, 'complete': not errors, 'errors': errors}

    def verify_sent(self, email, item_id=None, message_id=None, subject=None, expect_to=None,
                    expect_attachments=None, deep=False):
        from outlook_backend import MailError
        if not (item_id or message_id or subject):
            raise MailError('Give the draft ID from send, a sent message ID, --message-id or --subject.')
        session, binding = self.mailbox(email)
        store_id = str(binding[0].StoreID)
        journal = json.loads(self.journal_path.read_text('utf-8')) if self.journal_path.exists() else {}
        entry = journal.get(send_key(email, item_id)) if item_id else None
        expected = dict((entry or {}).get('fingerprint') or {})
        sent = self.default_folder(binding, 5)
        candidates, method, item = [], None, None
        if item_id and not entry:
            try:
                item = session.GetItemFromID(item_id, store_id)
                if str(item.Parent.StoreID) != store_id:
                    raise MailError('This message does not belong to the chosen mailbox.')
                method = 'item_id'
                message_id = message_id or self._prop(item, MSGID)
            except MailError:
                raise
            except Exception:
                raise MailError('Unknown ID: no send journal for this draft and no message with this ID.') from None
        if item is None:
            candidates, method = self._find_sent(session, store_id, sent, message_id or expected.get('message_id'),
                                                 subject or expected.get('subject'), expected.get('search_key'),
                                                 (entry or {}).get('requested_at'))
            if candidates:
                item = session.GetItemFromID(candidates[0], store_id)
        if expect_to:
            expected['recipients'] = normalize_addresses(expect_to)
        elif expected.get('recipients'):
            expected['recipients'] = [r['address'] for r in expected['recipients']]
        if subject:
            expected['subject'] = subject
        if expect_attachments is not None:
            expected['attachment_count'] = int(expect_attachments)
        if not entry or 'attachments' not in (entry.get('fingerprint') or {}):
            expected.pop('attachments', None)
        result = {'email': str(email).lower(), 'checked_at': iso(utcnow()), 'source': 'Outlook Classic',
                  'submitted': bool(entry and entry.get('submitted')) or None,
                  'submitted_at': (entry or {}).get('submitted_at'), 'journal_state': (entry or {}).get('state'),
                  'match_method': method, 'sent_candidates': len(candidates) if candidates else (1 if item else 0),
                  'in_sent_items': False}
        details = self._sent_details(item) if item is not None else {}
        if item is not None:
            in_sent = str(item.Parent.EntryID) == str(sent.EntryID)
            result.update(in_sent_items=in_sent, sent_item={'id': details['id'], 'folder_path': details['folder_path']},
                          message_id=details['message_id'], subject=details['subject'], sent_at=details['sent_at'],
                          size=details['size'], recipients=details['recipients'], attachments=details['attachments'],
                          attachment_count=len([a for a in details['attachments'] if not a['inline']]),
                          delivery_receipt_requested=details['delivery_receipt_requested'],
                          read_receipt_requested=details['read_receipt_requested'])
            if in_sent:
                result['submitted'] = True
        else:
            result.update(message_id=message_id or expected.get('message_id'), subject=subject or expected.get('subject'),
                          recipients=[{'address': a} for a in expected.get('recipients', [])], attachments=[])
        outbox = self._outbox(binding, result.get('subject'), expected.get('search_key') or details.get('search_key'),
                              result.get('message_id'))
        result['outbox_count'] = outbox['count']
        result['outbox_matching'] = outbox['matching']
        result['outbox_empty'] = not outbox['matching']
        result['expected'] = {k: v for k, v in expected.items() if k in ('subject', 'recipients', 'attachments', 'attachment_count')}
        result['checks'] = compare_expected(result['expected'], result) if item is not None else {}
        result['reports'], result['weak_reports'] = {}, []
        result['recipient_mailboxes'] = []
        if item is not None:
            since = result.get('sent_at') or (entry or {}).get('requested_at')
            reports, folders, errors = self._report_candidates(session, binding, since, deep=bool(deep))
            addresses = [r['address'] for r in result['recipients']]
            result['reports'], result['weak_reports'] = evaluate_reports(reports, result, addresses)
            result['bounce_scan'] = {'since': since, 'deep': bool(deep), 'folders': folders[:50],
                                     'scope': 'whole mailbox' if deep else 'Inbox with subfolders, Junk and Deleted Items',
                                     'folders_scanned': len(folders), 'report_candidates': len(reports), 'errors': errors}
            own = {a['email'].lower() for a in self.expected}
            for address in addresses:
                if address in own and result.get('message_id'):
                    try:
                        located = self.locate_message(address, result['message_id'])
                        result['recipient_mailboxes'].append({'recipient': address, 'backend': 'outlook', 'checked': True,
                                                              'found': located['found'], 'items': located['items'],
                                                              'complete': located['complete']})
                    except Exception as exc:
                        result['recipient_mailboxes'].append({'recipient': address, 'backend': 'outlook', 'checked': False,
                                                              'found': False, 'error': str(exc)[:200]})
        return finalize(result)
