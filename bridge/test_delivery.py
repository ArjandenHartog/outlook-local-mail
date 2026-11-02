"""verify_sent: matching, bounce/receipt classification, own-mailbox proof and routing. No real mail."""
import concurrent.futures
import datetime as dt
import json
import tempfile
import unittest
from email.message import EmailMessage
from email import policy
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch

import delivery
import imap_backend as adapter
import mail_index
from delivery import MSGID, SUBMIT, SEARCH_KEY, SIZE, RECEIVED, RECIPIENT_SMTP, ORIGINAL_SUBJECT, ORIGINAL_SUBMIT, SENDER_SMTP
from outlook_backend import OutlookBackend, MailError
from server import Manager
from test_manager import Collection

SENT_AT = dt.datetime(2026, 10, 20, 7, 45, 56, tzinfo=dt.timezone.utc)
MID = '<abc123@eurprd07.prod.outlook.com>'


class Props:
    def __init__(self, values):
        self.values = values

    def GetProperty(self, name):
        if name in self.values:
            return self.values[name]
        raise RuntimeError('property missing')


def folder(store, identifier, path, items=()):
    return NS(EntryID=identifier, FolderPath=path, StoreID=store, DefaultItemType=0, Folders=[], Items=list(items))


def recipient(address):
    return NS(Name=address, Type=1, Address=address, PropertyAccessor=Props({RECIPIENT_SMTP: address}))


class PureLogicTests(unittest.TestCase):
    target = {'message_id': MID, 'subject': 'Offerte', 'sent_at': SENT_AT.isoformat()}

    def test_classifies_exchange_classes_subjects_and_dsn_bodies(self):
        self.assertEqual(delivery.classify_report('REPORT.IPM.Note.NDR'), 'ndr')
        self.assertEqual(delivery.classify_report('REPORT.IPM.Note.DR'), 'delivery_receipt')
        self.assertEqual(delivery.classify_report('REPORT.IPM.Note.IPNRN'), 'read_receipt')
        self.assertEqual(delivery.classify_report('REPORT.IPM.Note.IPNNRN'), 'not_read')
        self.assertEqual(delivery.classify_report('', 'Onbestelbaar: Offerte'), 'ndr')
        self.assertEqual(delivery.classify_report('', 'Afgeleverd: Offerte'), 'delivery_receipt')
        self.assertEqual(delivery.classify_report('', 'Undelivered Mail', 'mailer-daemon@x', 'Action: failed'), 'ndr')
        self.assertEqual(delivery.classify_report('', 'Successful Mail Delivery Report', 'x', 'Action: delivered'), 'delivery_receipt')
        self.assertIsNone(delivery.classify_report('IPM.Note', 'Gewone mail', 'johan@example.test', 'Hallo'))

    def test_match_strength_prefers_message_id_then_original_subject_and_time(self):
        self.assertEqual(delivery.match_report({'body': 'Original-Message-ID: ' + MID.upper()}, self.target), 'message_id')
        report = {'original_subject': 'Offerte', 'original_submit': (SENT_AT + dt.timedelta(seconds=40)).isoformat()}
        self.assertEqual(delivery.match_report(report, self.target), 'original_subject_time')
        old = {'original_subject': 'Offerte', 'original_submit': (SENT_AT - dt.timedelta(days=3)).isoformat(),
               'subject': 'Undeliverable: Offerte'}
        self.assertEqual(delivery.match_report(old, self.target), 'subject_only')
        self.assertIsNone(delivery.match_report({'subject': 'Undeliverable: Iets anders', 'body': '<other@x>'}, self.target))

    def base(self, **extra):
        result = {'in_sent_items': True, 'outbox_empty': True, 'sent_at': SENT_AT.isoformat(), 'message_id': MID,
                  'subject': 'Offerte', 'recipients': [{'address': 'bob@example.test'}],
                  'bounce_scan': {'folders_scanned': 3, 'scope': 'Postvak IN'}, 'reports': {}, 'weak_reports': [],
                  'recipient_mailboxes': []}
        result.update(extra)
        return result

    def test_no_bounce_is_not_delivery_proof(self):
        result = delivery.finalize(self.base())
        self.assertEqual(result['state'], 'sent_no_bounce_seen')
        self.assertFalse(result['bounce_detected'])
        self.assertFalse(result['delivery_verified'])
        self.assertEqual(result['unverified_recipients'], ['bob@example.test'])
        self.assertIn('not proof of delivery', result['explanation'])

    def test_matched_ndr_is_bounce_and_names_failed_recipient(self):
        reports = [{'message_class': 'REPORT.IPM.Note.NDR', 'subject': 'Onbestelbaar: Offerte',
                    'body': 'Het bericht aan bob@example.test is niet afgeleverd. Message-ID: ' + MID}]
        found, weak = delivery.evaluate_reports(reports, self.base(), ['bob@example.test'])
        result = delivery.finalize(self.base(reports=found, weak_reports=weak))
        self.assertTrue(result['bounce_detected'])
        self.assertEqual(result['state'], 'bounced')
        self.assertEqual(result['failed_recipients'], ['bob@example.test'])
        self.assertFalse(result['delivery_verified'])

    def test_ndr_for_sibling_of_same_mail_merge_is_not_a_bounce_for_this_message(self):
        report = {'message_class': 'REPORT.IPM.Note.NDR', 'subject': 'Undeliverable: Offerte',
                  'original_subject': 'Offerte', 'original_submit': (SENT_AT + dt.timedelta(seconds=20)).isoformat(),
                  'body': 'Delivery to carol@example.test failed.'}
        found, weak = delivery.evaluate_reports([report], self.base(), ['bob@example.test'])
        self.assertEqual(found['ndr'], [])
        self.assertEqual(len(weak), 1)
        self.assertFalse(delivery.finalize(self.base(reports=found, weak_reports=weak))['bounce_detected'])
        report['body'] += ' Also bob@example.test failed.'
        found, _ = delivery.evaluate_reports([report], self.base(), ['bob@example.test'])
        self.assertEqual(len(found['ndr']), 1)

    def test_subject_only_ndr_is_only_a_possible_bounce(self):
        reports = [{'message_class': 'REPORT.IPM.Note.NDR', 'subject': 'Undeliverable: Offerte', 'body': 'no id'}]
        found, weak = delivery.evaluate_reports(reports, self.base(), ['bob@example.test'])
        result = delivery.finalize(self.base(reports=found, weak_reports=weak))
        self.assertFalse(result['bounce_detected'])
        self.assertEqual(len(result['possible_bounces']), 1)

    def test_delivery_receipt_and_own_mailbox_prove_delivery_per_recipient(self):
        recipients = [{'address': 'bob@example.test'}, {'address': 'me@example.test'}]
        receipt = {'message_class': 'REPORT.IPM.Note.DR', 'subject': 'Delivered: Offerte',
                   'original_subject': 'Offerte', 'original_submit': SENT_AT.isoformat(),
                   'body': 'Delivered to bob@example.test'}
        found, _ = delivery.evaluate_reports([receipt], self.base(), ['bob@example.test', 'me@example.test'])
        partial = delivery.finalize(self.base(recipients=recipients, reports=found))
        self.assertEqual(partial['state'], 'partially_verified')
        self.assertEqual(partial['unverified_recipients'], ['me@example.test'])
        mailboxes = [{'recipient': 'me@example.test', 'found': True}]
        result = delivery.finalize(self.base(recipients=recipients, reports=found, recipient_mailboxes=mailboxes))
        self.assertTrue(result['delivery_verified'])
        self.assertEqual(result['state'], 'delivered')
        self.assertEqual(result, delivery.finalize(result))  # idempotent for server-side re-finalizing

    def test_message_still_in_outbox(self):
        result = delivery.finalize(self.base(in_sent_items=False, outbox_empty=False))
        self.assertEqual(result['state'], 'in_outbox')
        self.assertFalse(result['delivery_verified'])

    def test_compare_expected_recipients_names_and_regular_attachment_count(self):
        actual = {'subject': 'Offerte', 'recipients': [{'address': 'BOB@example.test'}],
                  'attachments': [{'name': 'signature.png', 'size': 5, 'inline': True}, {'name': 'a.pdf', 'size': 102}]}
        checks = delivery.compare_expected({'subject': 'Offerte', 'recipients': ['bob@example.test', 'x@example.test'],
                                            'attachments': [{'name': 'signature.png', 'size': 5}, {'name': 'a.pdf', 'size': 100}],
                                            'attachment_count': 1}, actual)
        self.assertTrue(checks['subject'])
        self.assertFalse(checks['recipients'])
        self.assertEqual(checks['recipients_missing'], ['x@example.test'])
        self.assertTrue(checks['attachments'])
        self.assertTrue(checks['attachment_count'])
        self.assertTrue(checks['attachment_sizes_within_5_percent'])

    def test_ensure_indexed_requests_refresh_only_when_sent_copy_missing(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'index.sqlite3'
            writer = mail_index.SnapshotWriter(path)
            try:
                writer.add([dict(email='a@example.test', store_id='s', id='short-id', folder_id='f', folder_path='Verzonden items',
                                 subject='Offerte', sender='', sender_email='', received='2026-10-20T07:45:00+00:00', unread=0, kind='sent')])
                writer.publish({'complete': True, 'indexed_at': mail_index.utcnow(), 'invalidation': ''})
            finally:
                writer.close()
            schedule, invalidate = Mock(), Mock()
            present = delivery.ensure_indexed('a@example.test', 'Offerte', SENT_AT.isoformat(), path, schedule, invalidate)
            self.assertTrue(present['contains_sent_item'])
            schedule.assert_not_called()
            missing = delivery.ensure_indexed('a@example.test', 'Nieuw', SENT_AT.isoformat(), path, schedule, invalidate)
            self.assertFalse(missing['contains_sent_item'])
            self.assertTrue(missing['refresh_requested'])
            schedule.assert_called_once()
            invalidate.assert_called_once()


class OutlookVerifyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        (root / 'accounts.json').write_text(json.dumps([{'email': 'a@example.test'}, {'email': 'me@example.test'}]))
        self.sent_item = NS(EntryID='sent-long-id', Subject='Offerte', Class=43,
                            Recipients=[recipient('bob@example.test'), recipient('me@example.test')],
                            Attachments=[NS(FileName='a.pdf', Size=100, PropertyAccessor=Props({}))],
                            PropertyAccessor=Props({MSGID: MID, SUBMIT: SENT_AT, SEARCH_KEY: b'\x01\x02', SIZE: 1234}))
        self.ndr = NS(EntryID='ndr-1', Subject='Onbestelbaar: Offerte', MessageClass='REPORT.IPM.Note.NDR',
                      Body='Het bericht aan bob@example.test is niet afgeleverd.\r\nMessage-ID: ' + MID,
                      PropertyAccessor=Props({SENDER_SMTP: 'MicrosoftExchange@example.test', RECEIVED: SENT_AT + dt.timedelta(minutes=1),
                                              ORIGINAL_SUBJECT: 'Offerte', ORIGINAL_SUBMIT: SENT_AT}))
        self.old_ndr = NS(EntryID='ndr-old', Subject='Onbestelbaar: Offerte', MessageClass='REPORT.IPM.Note.NDR', Body=MID,
                          PropertyAccessor=Props({RECEIVED: SENT_AT - dt.timedelta(days=2)}))
        self.copy = NS(EntryID='copy-1', PropertyAccessor=Props({MSGID: MID, RECEIVED: SENT_AT}))
        self.folders = {}
        self.stores = {}
        for email, store in (('a@example.test', 'store-a'), ('me@example.test', 'store-me')):
            defaults = {n: folder(store, f'{store}-{n}', f'\\\\{email}\\f{n}') for n in (3, 4, 5, 6, 16, 23)}
            defaults[0] = folder(store, f'{store}-root', f'\\\\{email}')
            defaults[0].Folders = [defaults[n] for n in (3, 4, 5, 6, 16, 23)]
            self.folders[email] = defaults
            self.stores[email] = NS(StoreID=store, GetDefaultFolder=defaults.__getitem__, GetRootFolder=lambda d=defaults: d[0])
        a = self.folders['a@example.test']
        a[5].Items = [self.sent_item]
        a[6].Items = [self.ndr, self.old_ndr]
        self.folders['me@example.test'][6].Items = [self.copy]
        self.sent_item.Parent = a[5]
        items = {i.EntryID: i for i in (self.sent_item, self.ndr, self.old_ndr, self.copy)}
        accounts = Collection(NS(SmtpAddress=email, DeliveryStore=store) for email, store in self.stores.items())
        session = NS(Accounts=accounts, GetItemFromID=lambda identifier, store: items[identifier])
        self.backend = OutlookBackend(root / 'accounts.json', NS(Session=session), root / 'send-journal.json')
        self.backend._table = self.table

    def tearDown(self):
        self.temp.cleanup()

    @staticmethod
    def table(folder, restriction, columns):
        rows = []
        for item in folder.Items:
            props = item.PropertyAccessor.values
            if f'"{MSGID}" =' in restriction and props.get(MSGID) != MID:
                continue
            if 'httpmail:subject" =' in restriction and getattr(item, 'Subject', None) != 'Offerte':
                continue
            if 'REPORT.%' in restriction and not str(getattr(item, 'MessageClass', '')).startswith('REPORT.'):
                continue
            values = {'EntryID': item.EntryID, 'MessageClass': getattr(item, 'MessageClass', 'IPM.Note'),
                      'Subject': getattr(item, 'Subject', '')}
            rows.append([values.get(c, props.get(c)) for c in columns])
        return rows

    def journal(self):
        key = delivery.send_key('a@example.test', 'draft-1')
        fingerprint = {'subject': 'Offerte', 'search_key': '0102', 'recipients': [{'address': 'bob@example.test'}, {'address': 'me@example.test'}],
                       'attachments': [{'name': 'a.pdf', 'size': 100}]}
        self.backend.journal_path.write_text(json.dumps({key: {'state': 'submitted', 'submitted': True,
            'requested_at': SENT_AT.isoformat(), 'submitted_at': SENT_AT.isoformat(), 'fingerprint': fingerprint}}))

    def test_draft_id_uses_journal_search_key_and_detects_bounce_and_own_arrival(self):
        self.journal()
        result = self.backend.verify_sent('a@example.test', item_id='draft-1')
        self.assertEqual(result['match_method'], 'search_key')
        self.assertTrue(result['in_sent_items'])
        self.assertTrue(result['outbox_empty'])
        self.assertEqual(result['message_id'], MID)
        self.assertEqual(result['sent_at'], SENT_AT.isoformat())
        self.assertTrue(all(result['checks'][k] for k in ('subject', 'recipients', 'attachments')))
        self.assertTrue(result['bounce_detected'])
        self.assertEqual([b['id'] for b in result['bounces']], ['ndr-1'])  # old NDR outside the time window
        self.assertEqual(result['failed_recipients'], ['bob@example.test'])
        self.assertEqual(result['delivered_recipients'], [{'recipient': 'me@example.test', 'evidence': ['own_mailbox']}])
        self.assertFalse(result['delivery_verified'])
        self.assertEqual(result['state'], 'bounced')

    def test_message_id_lookup_without_journal_and_outbox_pending(self):
        self.folders['a@example.test'][6].Items = []
        self.folders['a@example.test'][4].Items = [NS(EntryID='out-1', Subject='Offerte', PropertyAccessor=Props({SEARCH_KEY: b'\x01\x02'}))]
        result = self.backend.verify_sent('a@example.test', message_id=MID, expect_to='bob@example.test,me@example.test')
        self.assertEqual(result['match_method'], 'message_id')
        self.assertFalse(result['outbox_empty'])
        self.assertFalse(result['bounce_detected'])
        self.assertTrue(result['checks']['recipients'])

    def test_requires_a_reference_and_foreign_ids_are_rejected(self):
        with self.assertRaises(MailError):
            self.backend.verify_sent('a@example.test')
        with self.assertRaises(MailError):
            self.backend.verify_sent('a@example.test', item_id='copy-1')  # copy lives in another store

    def test_send_records_fingerprint_without_body(self):
        draft = NS(EntryID='draft-9', Subject='Offerte', Sent=False, Parent=self.folders['a@example.test'][16],
                   Recipients=Mock(Count=1, ResolveAll=Mock(return_value=True), __iter__=lambda s: iter([recipient('bob@example.test')])),
                   Attachments=[], Body='geheim', PropertyAccessor=Props({SEARCH_KEY: b'\xaa'}),
                   OriginatorDeliveryReportRequested=True, ReadReceiptRequested=False, Save=Mock(), Send=Mock())
        account = self.backend.session().Accounts[0]
        draft.SendUsingAccount = account
        with patch.object(self.backend, 'message', return_value=(draft, (self.stores['a@example.test'], account, None))), \
                patch.object(OutlookBackend, 'set_sender'), patch.object(self.backend, 'invalidate_index'):
            receipt = self.backend.send('a@example.test', 'draft-9')
        self.assertEqual(receipt['fingerprint']['search_key'], 'aa')
        self.assertEqual(receipt['fingerprint']['recipients'][0]['address'], 'bob@example.test')
        self.assertTrue(receipt['fingerprint']['delivery_receipt_requested'])
        self.assertNotIn('geheim', self.backend.journal_path.read_text('utf-8'))


class ReceiptRequestTests(unittest.TestCase):
    def test_outlook_draft_sets_receipt_flags_and_rejects_non_bool(self):
        from test_manager import BackendTests
        case = BackendTests('test_new_draft_is_saved_in_target_account_unsent')
        case.setUp()
        try:
            result = case.backend.draft('a@example.test', subject='T', body='B', request_delivery_receipt=True, request_read_receipt=True)
            draft = case.drafts.Items[0]
            self.assertTrue(draft.OriginatorDeliveryReportRequested and draft.ReadReceiptRequested)
            self.assertTrue(result['delivery_receipt_requested'] and result['read_receipt_requested'])
            with self.assertRaises(MailError):
                case.backend.draft('a@example.test', request_read_receipt='yes')
        finally:
            case.tearDown()


class ImapVerifyTests(unittest.TestCase):
    def setUp(self):
        from test_imap_backend import ImapTests
        self.case = ImapTests('test_send_logs_in_primary_uses_alias_envelope_and_is_durable')
        self.case.setUp()
        self.backend, self.conn, self.smtp = self.case.backend, self.case.conn, self.case.smtp
        self.email = self.case.config['email']

    def tearDown(self):
        self.case.tearDown()

    def test_dsn_and_read_receipt_requests_reach_smtp_and_marker_is_stripped(self):
        self.smtp.has_extn.return_value = True
        draft = self.backend.draft(self.email, body='Test', to='bob@example.test', request_delivery_receipt=True, request_read_receipt=True)
        raw = self.conn.data['Drafts'][adapter.decode(draft['id'], self.email, True)['u']]
        self.assertIn(b'Disposition-Notification-To', raw)
        self.assertIn(adapter.DSN_MARKER.encode(), raw)
        receipt = self.backend.send(self.email, draft['id'])
        self.assertTrue(receipt['dsn_supported'])
        self.assertEqual(self.smtp.sendmail.call_args.kwargs['rcpt_options'], ['NOTIFY=SUCCESS,FAILURE,DELAY'])
        self.assertNotIn(adapter.DSN_MARKER.encode(), self.smtp.sendmail.call_args.args[2])
        self.assertTrue(receipt['fingerprint']['message_id'].startswith('<'))

    def test_verify_finds_sent_copy_and_dsn_bounce_by_message_id(self):
        draft = self.backend.draft(self.email, subject='Offerte', body='Test', to='bob@example.test')
        receipt = self.backend.send(self.email, draft['id'])
        message_id = receipt['fingerprint']['message_id']
        bounce = EmailMessage(policy=policy.SMTP)
        bounce['From'] = 'MAILER-DAEMON@mail.example.test'
        bounce['Subject'] = 'Undelivered Mail Returned to Sender'
        bounce['Date'] = 'Tue, 20 Oct 2026 08:00:00 +0000'
        bounce.set_content('Final-Recipient: rfc822; bob@example.test\nAction: failed\nMessage-ID: ' + message_id)
        self.conn.data['INBOX'][50] = bounce.as_bytes()
        before = len(self.conn.calls)
        result = self.backend.verify_sent(self.email, item_id=draft['id'])
        self.assertTrue(result['in_sent_items'])
        self.assertEqual(result['message_id'], message_id)
        self.assertTrue(result['bounce_detected'])
        self.assertEqual(result['failed_recipients'], ['bob@example.test'])
        self.assertTrue(result['checks']['recipients'])
        self.assertFalse(any(c[0] in ('store', 'expunge', 'move') for c in self.conn.calls[before:]))


class ServerRoutingTests(unittest.TestCase):
    def test_verify_sent_cross_checks_imap_recipient_and_index(self):
        worker = Mock()
        base = {'email': 'a@example.test', 'in_sent_items': True, 'outbox_empty': True, 'message_id': MID, 'subject': 'Offerte',
                'sent_at': SENT_AT.isoformat(), 'recipients': [{'address': 'me@example.com'}], 'reports': {}, 'weak_reports': [],
                'recipient_mailboxes': [], 'sent_item': {'id': 'x'}}
        future = concurrent.futures.Future()
        future.set_result(base)
        worker.submit.return_value = future
        server = Manager(('127.0.0.1', 0), worker)
        try:
            server.imap.configs = {'me@example.com': {}}
            server.imap.locate_message = Mock(return_value={'found': True, 'items': [{'id': 'imap:1'}], 'complete': True})
            with patch('mail_index.allowed_accounts', return_value={'a@example.test', 'me@example.com'}), \
                    patch('delivery.ensure_indexed', return_value={'contains_sent_item': True}) as indexed:
                result = server.verify_sent({'email': 'a@example.test', 'message_id': MID})
            self.assertTrue(result['delivery_verified'])
            server.imap.locate_message.assert_called_once_with('me@example.com', MID)
            indexed.assert_called_once()
            with self.assertRaises(MailError):
                server.job('verify-000000000001', 'verify_sent', {'email': 'a@example.test', 'body': 'x'})
        finally:
            server.server_close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
