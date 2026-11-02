import concurrent.futures
import datetime as dt
import json
from pathlib import Path
import tempfile
import threading
from types import SimpleNamespace as NS
import unittest
from unittest.mock import patch
from urllib.request import Request, urlopen
from urllib.error import HTTPError

from outlook_backend import OutlookBackend, MailError, received_timestamp
from server import Manager


class Collection(list):
    @property
    def Count(self):
        return len(self)

    def Item(self, index):
        return self[index - 1]

    def Sort(self, field, descending):
        self.sort(key=lambda x: x.ReceivedTime, reverse=descending)

    def Restrict(self, query):
        self.restriction = query
        return self

    def Add(self, kind):
        item = Message('new-draft', self.folder)
        self.append(item)
        return item


class Message:
    def __init__(self, identifier, folder):
        self.EntryID = identifier
        self.Parent = folder
        self.Class = 43
        self.Subject = '<script>danger()</script>'
        self.Body = 'Voorbeeldbericht'
        self.To = self.CC = ''
        self.SenderName = 'Afzender'
        self.SenderEmailAddress = 'sender@example.test'
        self.ReceivedTime = dt.datetime.now(dt.timezone.utc)
        self.UnRead = True
        self.Attachments = Collection()
        self.Sent = False
        self.saves = 0

    def Save(self):
        self.saves += 1

    def Send(self):
        raise AssertionError('No test operation may send mail')

    def Reply(self):
        reply = Message('reply-draft', self.Parent)
        reply.Body = 'Geciteerd origineel'
        reply.To = self.SenderEmailAddress
        return reply

    def Move(self, folder):
        self.Parent = folder
        folder.Items.append(self)
        return self


def fake_folder(store_id, identifier, name):
    folder = NS(StoreID=store_id, EntryID=identifier, Name=name, FolderPath=name,
                Items=Collection(), Folders=Collection(), UnReadItemCount=0)
    folder.Items.folder = folder
    return folder


class BackendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / 'accounts.json'
        self.path.write_text(json.dumps([{'email': 'a@example.test'}, {'email': 'b@example.test'}]))
        self.inbox = fake_folder('store-a', 'inbox-a', 'Postvak IN')
        self.drafts = fake_folder('store-a', 'drafts-a', 'Concepten')
        self.child = fake_folder('store-a', 'sub-a', 'Submap')
        self.inbox.Folders.append(self.child)
        self.item = Message('mail-a', self.inbox)
        self.inbox.Items.extend([self.item, NS(Class=26, ReceivedTime=self.item.ReceivedTime)])
        self.store = NS(StoreID='store-a', DisplayName='Account A', GetDefaultFolder=lambda n: {6: self.inbox, 16: self.drafts}[n], GetRootFolder=lambda: self.inbox)
        self.account = NS(SmtpAddress='A@example.test', DeliveryStore=self.store)
        self.foreign = Message('mail-b', fake_folder('store-b', 'inbox-b', 'Foreign'))
        self.session = NS(Accounts=Collection([self.account]), Stores=Collection([self.store]), Offline=False,
                          GetItemFromID=lambda identifier, store: self.foreign if identifier == 'mail-b' else self.item,
                          GetFolderFromID=lambda identifier, store: self.foreign.Parent if identifier == 'inbox-b' else self.child,
                          SyncObjects=Collection())
        self.backend = OutlookBackend(self.path, NS(Session=self.session))

    def tearDown(self):
        self.temp.cleanup()

    def test_status_matches_email_case_and_missing_account(self):
        state = self.backend.status()
        self.assertEqual(state['connected'], 1)
        self.assertFalse(state['accounts'][1]['connected'])
        self.assertFalse(state['sync_verified'])

    def test_implicit_primary_sender_requires_matching_default_and_first_account_store(self):
        self.item.SendUsingAccount = None
        self.item.Session = NS(Accounts=Collection([self.account]), DefaultStore=self.store)
        self.assertEqual(OutlookBackend.sender_email(self.item), 'A@example.test')

    def test_implicit_primary_sender_never_applies_to_other_mailbox(self):
        self.foreign.SendUsingAccount = None
        self.foreign.Session = NS(Accounts=Collection([self.account]), DefaultStore=self.store)
        self.assertEqual(OutlookBackend.sender_email(self.foreign), '')

    def test_received_time_uses_utc_delivery_property_instead_of_mislabeled_local_date(self):
        delivered = dt.datetime(2026, 10, 5, 8, 2, tzinfo=dt.timezone.utc)
        self.item.ReceivedTime = delivered + dt.timedelta(hours=2)
        self.item.PropertyAccessor = NS(GetProperty=lambda name: delivered)
        self.assertEqual(received_timestamp(self.item), '2026-10-05T08:02:00+00:00')

    def test_received_time_converts_local_date_when_delivery_property_missing(self):
        delivered = dt.datetime(2026, 10, 5, 8, 2, tzinfo=dt.timezone.utc)
        def missing(name):
            raise RuntimeError('missing')
        self.item.PropertyAccessor = NS(GetProperty=missing, LocalTimeToUTC=lambda value: delivered)
        self.assertEqual(received_timestamp(self.item), '2026-10-05T08:02:00+00:00')

    def test_background_starts_classic_when_no_active_object(self):
        self.backend.client = None
        explorer = NS(WindowState=2)
        app = NS(Session=self.session, Explorers=[explorer])
        with patch('win32com.client.GetActiveObject', side_effect=[RuntimeError('closed'), app]), patch('subprocess.Popen') as start, patch('pathlib.Path.is_file', return_value=True):
            self.assertIs(self.backend.session(), self.session)
            self.assertEqual(start.call_count, 1)
            self.assertEqual(start.call_args.args[0][-1], '/recycle')
            self.assertEqual(explorer.WindowState, 1)

    def test_background_start_wait_is_bounded_and_launch_is_throttled(self):
        self.backend.client = None
        with patch('win32com.client.GetActiveObject', side_effect=RuntimeError('closed')), patch('subprocess.Popen') as start, patch('pathlib.Path.is_file', return_value=True), patch('time.sleep'):
            for _ in range(2):
                with self.assertRaisesRegex(MailError, 'not available yet'):
                    self.backend.session()
            self.assertEqual(start.call_count, 1)

    def test_lost_com_connection_is_recovered_on_next_check(self):
        self.backend.client = NS()
        with self.assertRaises(MailError):
            self.backend.session()
        self.assertIsNone(self.backend.client)
        with patch('win32com.client.GetActiveObject', return_value=NS(Session=self.session)):
            self.assertIs(self.backend.session(), self.session)

    def test_mailbox_scope_rejects_foreign_message_and_folder(self):
        with self.assertRaises(MailError):
            self.backend.read('a@example.test', 'mail-b')
        with self.assertRaises(MailError):
            self.backend.messages('a@example.test', 'inbox-b')
        with self.assertRaises(MailError):
            self.backend.read('unknown@example.test', 'mail-a')

    def test_read_does_not_mark_message_read(self):
        result = self.backend.read('a@example.test', 'mail-a')
        self.assertEqual(result['body'], 'Voorbeeldbericht')
        self.assertTrue(self.item.UnRead)

    def test_subfolders_and_pagination_non_mail(self):
        self.assertIn('sub-a', [f['id'] for f in self.backend.folders('a@example.test')['folders']])
        first = self.backend.messages('a@example.test', limit=1)
        self.assertEqual(first['next_offset'], 1)
        second = self.backend.messages('a@example.test', offset=1, limit=1)
        self.assertEqual(second['items'], [])
        self.assertIsNone(second['next_offset'])

    def test_search_literal_is_escaped_and_page_limit_enforced(self):
        self.backend.messages('a@example.test', query="O'Brien")
        self.assertIn("O''Brien", self.inbox.Items.restriction)
        with self.assertRaises(MailError):
            self.backend.messages('a@example.test', limit=201)

    def test_new_draft_is_saved_in_target_account_unsent(self):
        result = self.backend.draft('a@example.test', subject='Test', body='Inhoud')
        draft = self.drafts.Items[0]
        self.assertEqual(draft.Subject, 'Test')
        self.assertEqual(draft.To, '')
        self.assertIs(draft.SendUsingAccount, self.account)
        self.assertEqual(draft.Parent.StoreID, 'store-a')
        self.assertFalse(result['sent'])
        self.assertGreaterEqual(draft.saves, 1)

    def test_com_sender_uses_interface_reference_setter(self):
        from unittest.mock import Mock
        import pythoncom
        dispatch = Mock()
        dispatch.GetIDsOfNames.return_value = 64209
        draft = NS(_oleobj_=dispatch, PropertyAccessor=Mock(), SendUsingAccount=self.account)
        self.account.CurrentUser = NS(AddressEntry=NS(Address='a@example.test'))
        self.backend.set_sender(draft, self.account)
        dispatch.Invoke.assert_called_once_with(64209, 0, pythoncom.DISPATCH_PROPERTYPUTREF, 0, self.account)
        dispatch.GetIDsOfNames.assert_called_once_with('SendUsingAccount')
        draft.PropertyAccessor.SetProperty.assert_not_called()

    def test_reply_preserves_quote_recipient_and_target_drafts(self):
        self.backend.draft('a@example.test', body='Mijn antwoord', reply_id='mail-a')
        reply = self.drafts.Items[0]
        self.assertEqual(reply.To, 'sender@example.test')
        self.assertIn('Geciteerd origineel', reply.Body)
        self.assertTrue(reply.Body.startswith('Mijn antwoord'))
        self.assertIs(reply.Parent, self.drafts)

    def test_sync_request_is_not_completion(self):
        result = self.backend.sync()
        self.assertFalse(result['completed'])
        self.assertFalse(self.backend.status()['sync_verified'])

    def sendable(self):
        from unittest.mock import Mock
        self.item.Parent = self.drafts
        self.item.Recipients = NS(Count=1, ResolveAll=lambda:True)
        self.item.Send = Mock()
        self.backend.journal_path = self.path.parent / 'send-journal.json'

    def test_send_uses_selected_account_and_persists_duplicate_protection(self):
        self.sendable()
        receipt = self.backend.send('a@example.test','mail-a')
        self.assertTrue(receipt['submitted'])
        self.assertFalse(receipt['delivery_verified'])
        self.assertIs(self.item.SendUsingAccount,self.account)
        self.backend.send('a@example.test','mail-a')
        self.item.Send.assert_called_once()
        restarted = OutlookBackend(self.path, NS(Session=self.session), self.backend.journal_path)
        self.assertTrue(restarted.send('a@example.test','mail-a')['submitted'])
        self.item.Send.assert_called_once()

    def test_send_rejects_non_draft_and_unresolved_recipients(self):
        self.sendable()
        self.item.Parent = self.inbox
        with self.assertRaises(MailError):
            self.backend.send('a@example.test','mail-a')
        self.item.Parent = self.drafts
        self.item.Recipients.ResolveAll = lambda:False
        with self.assertRaises(MailError):
            self.backend.send('a@example.test','mail-a')
        self.item.Send.assert_not_called()

    def test_send_never_overwrites_sender_after_selecting_transport_account(self):
        from unittest.mock import Mock
        self.sendable()
        dispatch = Mock()
        dispatch.GetIDsOfNames.side_effect = lambda name: 1 if name == 'SendUsingAccount' else 2
        changes = []
        def invoke(property_id, *args):
            changes.append(property_id)
            # This models Outlook's observed behavior: assigning Sender clears
            # SendUsingAccount and leaves the submitted message stuck in Outbox.
            self.item.SendUsingAccount = self.account if property_id == 1 else None
        dispatch.Invoke.side_effect = invoke
        self.item._oleobj_ = dispatch
        self.account.CurrentUser = NS(AddressEntry=NS(Address='a@example.test'))
        self.item.Send.side_effect = lambda: self.assertIs(self.item.SendUsingAccount, self.account)
        self.assertTrue(self.backend.send('a@example.test', 'mail-a')['submitted'])
        self.assertEqual(changes, [1, 1])

    def test_send_exception_is_not_automatically_retried(self):
        self.sendable()
        self.item.Send.side_effect = RuntimeError('ambiguous transport result')
        for _ in range(2):
            with self.assertRaises(MailError):
                self.backend.send('a@example.test','mail-a')
        self.item.Send.assert_called_once()


class FakeWorker:
    def __init__(self):
        self.calls = []

    def submit(self, action, args):
        self.calls.append((action, args))
        future = concurrent.futures.Future()
        future.set_result({'ok': True, 'sent': False})
        return future


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.worker = FakeWorker()
        cls.server = Manager(('127.0.0.1', 0), cls.worker)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def request(self, action='status', headers=None, code='request-000000000001', args=None):
        h = {'Content-Type': 'application/json', 'X-Local-Token': self.server.token}
        h.update(headers or {})
        req = Request(self.server.origin + '/api/call', json.dumps({'action': action, 'args': args or {}, 'request_id': code}).encode(), h)
        try:
            with urlopen(req) as res:
                return res.status, json.load(res)
        except HTTPError as exc:
            return exc.code, json.load(exc)

    def test_missing_token_cross_origin_and_rebinding_are_denied(self):
        self.assertEqual(self.request(headers={'X-Local-Token': ''})[0], 401)
        self.assertEqual(self.request(headers={'Origin': 'https://other.example'})[0], 403)
        self.assertEqual(self.request(headers={'Host': 'attacker.example'})[0], 403)

    def test_unsupported_action_and_delete_never_enter_worker(self):
        before = len(self.worker.calls)
        self.assertEqual(self.request('unsupported')[0], 400)
        self.assertEqual(self.request('delete')[0], 400)
        self.assertEqual(len(self.worker.calls), before)

    def test_duplicate_draft_request_runs_once(self):
        before = len(self.worker.calls)
        for _ in range(2):
            self.assertEqual(self.request('draft', code='draft-000000000001', args={'email': 'a@example.test'})[0], 200)
        self.assertEqual(len(self.worker.calls), before + 1)
        self.assertEqual(self.request('draft', code='draft-000000000001', args={'email': 'b@example.test'})[0], 400)

    def test_second_listener_cannot_replace_session_token(self):
        with self.assertRaises(OSError):
            Manager(('127.0.0.1', self.server.server_port), FakeWorker())


if __name__ == '__main__':
    unittest.main(verbosity=2)
