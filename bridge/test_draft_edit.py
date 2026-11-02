"""update_draft / delete_draft: Outlook COM and IMAP. No test may send mail."""
import json
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch
from email import policy
from email.parser import BytesParser

import draft_edit
import imap_backend as adapter
import test_imap_backend as imap_fixtures
import test_manager as fixtures
from outlook_backend import MailError

HIDDEN = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'


class Att:
    def __init__(self, name, size, hidden=False):
        self.FileName, self.Size = name, size
        def get(prop):
            # Like a reopened Outlook draft: hidden=False, but the signature image keeps a content-id.
            if hidden and prop == 'http://schemas.microsoft.com/mapi/proptag/0x3712001F':
                return 'outlook-local-signature-abc'
            if prop == 'http://schemas.microsoft.com/mapi/proptag/0x3712001F':
                return ''
            raise KeyError(prop)
        self.PropertyAccessor = NS(GetProperty=get)


class Atts(fixtures.Collection):
    def Add(self, path, kind):
        self.append(Att(Path(path).name, Path(path).stat().st_size))

    def Remove(self, index):
        del self[index - 1]


SIGNATURE = '<div data-outlook-local-signature="true"><p>Groet, Jane</p></div>'
QUOTE = '<br><div id="quote">Oorspronkelijk bericht</div>'


class NativeDraftEditTests(unittest.TestCase):
    def setUp(self):
        self.f = fixtures.BackendTests(); self.f.setUp()
        self.backend, self.session = self.f.backend, self.f.session
        self.items = {'mail-a': self.f.item, 'mail-b': self.f.foreign}
        self.session.GetItemFromID = lambda identifier, store: self.items[identifier]
        self.backend.journal_path = Path(self.f.temp.name) / 'send-journal.json'
        self.draft = self.make('draft-1')

    def tearDown(self):
        self.f.tearDown()

    def make(self, identifier, **values):
        item = fixtures.Message(identifier, self.f.drafts)
        item.Subject, item.Body, item.To, item.CC, item.BCC = 'Oud onderwerp', 'Oude tekst', 'x@example.test', 'cc@example.test', ''
        item.Attachments = Atts([Att('rapport.pdf', 10), Att('signature.png', 5, hidden=True)])
        item.SendUsingAccount = self.f.account
        item.HTMLBody = ('<html><body><div data-outlook-local-body="true" style="font-size:11pt">Oude tekst</div><br>'
                         + SIGNATURE + QUOTE + '</body></html>')
        item.Delete = Mock()
        for key, value in values.items():
            setattr(item, key, value)
        self.f.drafts.Items.append(item); self.items[identifier] = item
        return item

    def update(self, **fields):
        return self.backend.update_draft('a@example.test', 'draft-1', **fields)

    def test_only_given_fields_change_and_nothing_is_sent(self):
        result = self.update(subject='Nieuw', to='y@example.test', bcc='b@example.test')
        self.assertEqual((self.draft.Subject, self.draft.To, self.draft.CC, self.draft.BCC),
                         ('Nieuw', 'y@example.test', 'cc@example.test', 'b@example.test'))
        self.assertIn('Oude tekst', self.draft.HTMLBody); self.assertEqual(len(self.draft.Attachments), 2)
        self.assertFalse(result['sent']); self.assertEqual(result['id'], 'draft-1'); self.assertFalse(result['id_changed'])
        self.assertEqual(result['changed'], ['bcc', 'subject', 'to']); self.assertEqual(self.draft.saves, 1)
        self.assertEqual(result['attachments'], [{'name': 'rapport.pdf', 'size': 10}])

    def test_empty_string_clears_field_but_omitted_stays(self):
        self.update(cc='')
        self.assertEqual(self.draft.CC, ''); self.assertEqual(self.draft.To, 'x@example.test')

    def test_body_replaces_authored_text_only_and_escapes_html(self):
        self.update(body='Nieuw <b>\nregel')
        html = self.draft.HTMLBody
        self.assertIn('>Nieuw &lt;b&gt;<br>regel</div>', html); self.assertNotIn('Oude tekst', html)
        self.assertTrue(html.endswith(SIGNATURE + QUOTE + '</body></html>'))

    def test_body_without_recognizable_structure_is_refused_without_partial_change(self):
        self.draft.HTMLBody = '<html><body>Handgeschreven antwoord</body></html>'
        with self.assertRaisesRegex(MailError, 'cannot be delimited'):
            self.update(subject='Nieuw', body='x')
        self.assertEqual(self.draft.Subject, 'Oud onderwerp'); self.assertEqual(self.draft.saves, 0)

    def test_attachment_remove_add_keeps_hidden_signature_image(self):
        new = Path(self.f.temp.name) / 'nieuw.txt'; new.write_bytes(b'abc')
        result = self.update(attachments_remove=['RAPPORT.pdf'], attachments_add=[str(new)])
        self.assertEqual([a.FileName for a in self.draft.Attachments], ['signature.png', 'nieuw.txt'])
        self.assertEqual(result['attachments'], [{'name': 'nieuw.txt', 'size': 3}])

    def test_attachment_errors_leave_draft_untouched(self):
        new = Path(self.f.temp.name) / 'rapport.pdf'; new.write_bytes(b'abc')
        for fields in ({'attachments_remove': ['bestaat-niet.pdf']}, {'attachments_remove': ['signature.png']},
                       {'attachments_add': [str(new)]}, {'attachments_add': [str(new.with_name('weg.pdf'))]}):
            with self.assertRaises((MailError, ValueError)):
                self.update(subject='Nieuw', **fields)
        self.assertEqual(self.draft.Subject, 'Oud onderwerp'); self.assertEqual(len(self.draft.Attachments), 2)
        self.assertEqual(self.draft.saves, 0)

    def test_unknown_and_empty_updates_refused(self):
        for fields in ({'reply_id': 'x'}, {'signature': 'none'}, {'subject': None}, {}, {'attachments_add': []},
                       {'subject': 'a\r\nBcc: evil@example.test'}):
            with self.assertRaises((MailError, ValueError)):
                self.update(**fields)
        self.assertEqual(self.draft.saves, 0)

    def test_only_unsent_drafts_in_own_drafts_folder(self):
        sent = self.make('sent-1', Sent=True)
        journal = self.make('journal-1')
        self.backend.journal_path.write_text(json.dumps({draft_edit_key('a@example.test', 'journal-1'): {'state': 'submitted'}}))
        for identifier in ('mail-a', 'sent-1', 'journal-1', 'mail-b'):
            with self.assertRaises(MailError):
                self.backend.update_draft('a@example.test', identifier, subject='Nieuw')
            with self.assertRaises(MailError):
                self.backend.delete_draft('a@example.test', identifier, reason='dubbel')
        self.assertEqual(self.f.item.saves, 0); sent.Delete.assert_not_called(); journal.Delete.assert_not_called()

    def test_delete_requires_reason_and_matching_subject_and_is_reversible(self):
        with self.assertRaises(MailError):
            self.backend.delete_draft('a@example.test', 'draft-1', reason=' ')
        with self.assertRaisesRegex(MailError, 'expect-subject'):
            self.backend.delete_draft('a@example.test', 'draft-1', reason='dubbel', expect_subject='Ander')
        self.draft.Delete.assert_not_called()
        result = self.backend.delete_draft('a@example.test', 'draft-1', reason='dubbel', expect_subject='Oud onderwerp')
        self.draft.Delete.assert_called_once_with()
        self.assertTrue(result['deleted']); self.assertFalse(result['permanent_delete'])


def draft_edit_key(email, item_id):
    import delivery
    return delivery.send_key(email, item_id)


class ImapDraftEditTests(unittest.TestCase):
    def setUp(self):
        self.f = imap_fixtures.ImapTests(); self.f.setUp()
        self.backend, self.conn, self.email = self.f.backend, self.f.conn, self.f.config['email']
        self.created = self.backend.draft(self.email, subject='Onderwerp', body='Eerste\ntekst', to='x@example.test')

    def tearDown(self):
        self.f.tearDown()

    def parse(self, item_id):
        uid = adapter.decode(item_id, self.email, True)['u']
        return BytesParser(policy=policy.default).parsebytes(self.conn.data['Drafts'][uid])

    def update(self, item_id=None, **fields):
        return self.backend.update_draft(self.email, item_id or self.created['id'], **fields)

    def test_body_replaced_keeps_signature_headers_and_leaves_one_draft(self):
        before = self.parse(self.created['id'])
        result = self.update(body='Nieuwe <tekst>\nmet regel', subject='Nieuw onderwerp')
        self.assertFalse(result['sent']); self.assertTrue(result['old_removed']); self.assertTrue(result['uid_changed'])
        self.assertFalse(result['uidvalidity_changed']); self.assertNotEqual(result['id'], self.created['id'])
        self.assertEqual(len(self.conn.data['Drafts']), 1)
        message = self.parse(result['id'])
        self.assertEqual(message['Message-ID'], before['Message-ID']); self.assertEqual(message['Subject'], 'Nieuw onderwerp')
        self.assertEqual(message['To'], 'x@example.test')
        plain = message.get_body(preferencelist=('plain',)).get_content().replace('\r\n', '\n')
        rich = message.get_body(preferencelist=('html',)).get_content()
        self.assertTrue(plain.startswith('Nieuwe <tekst>\nmet regel\n\nJane Doe me@example.com'))
        self.assertIn('Nieuwe &lt;tekst&gt;<br>met regel</div>', rich); self.assertNotIn('Eerste', rich + plain)
        self.assertIn('cid:signature-', rich); self.assertTrue(any(p.get_content_disposition() == 'inline' for p in message.walk()))
        self.assertEqual(self.backend.read(self.email, result['id'])['is_draft'], True)
        self.f.smtp.sendmail.assert_not_called()

    def test_reply_link_and_quoted_thread_survive_body_edit(self):
        reply = self.backend.draft(self.email, body='Antwoord', reply_id=self.f.item)
        result = self.update(reply['id'], body='Beter antwoord')
        message = self.parse(result['id'])
        self.assertEqual(message['In-Reply-To'], '<original@example.test>')
        self.assertIn('Hello there', message.get_body(preferencelist=('plain',)).get_content())
        self.assertIn('Hello there', message.get_body(preferencelist=('html',)).get_content())
        self.assertTrue(message.get_body(preferencelist=('plain',)).get_content().startswith('Beter antwoord'))

    def test_recipients_set_clear_and_bcc(self):
        result = self.update(to='a@example.test, b@example.test', cc='c@example.test', bcc='d@example.test')
        message = self.parse(result['id'])
        self.assertEqual((message['To'], message['Cc'], message['Bcc']), ('a@example.test, b@example.test', 'c@example.test', 'd@example.test'))
        self.assertEqual(self.backend.read(self.email, result['id'])['bcc'], 'd@example.test')
        cleared = self.parse(self.update(result['id'], cc='')['id'])
        self.assertIsNone(cleared['Cc']); self.assertEqual(cleared['Bcc'], 'd@example.test')

    def test_attachments_add_remove_and_errors_do_not_touch_draft(self):
        report, other = self.f.root / 'rapport.pdf', self.f.root / 'ander.txt'
        report.write_bytes(b'%PDF'); other.write_bytes(b'tekst')
        withfile = self.backend.draft(self.email, body='Ha', attachments=[str(report)])
        appends = len([c for c in self.conn.calls if c[0] == 'append'])
        for fields in ({'attachments_remove': ['weg.pdf']}, {'attachments_remove': ['signature.png']}, {'attachments_add': [str(report)]}):
            with self.assertRaises(ValueError):
                self.update(withfile['id'], **fields)
        self.assertEqual(len([c for c in self.conn.calls if c[0] == 'append']), appends)
        result = self.update(withfile['id'], attachments_remove=['RAPPORT.pdf'], attachments_add=[str(other)])
        message = self.parse(result['id'])
        self.assertEqual([p.get_filename() for p in message.iter_attachments() if p.get_content_disposition() == 'attachment'], ['ander.txt'])
        self.assertTrue(any(p.get_content_disposition() == 'inline' for p in message.walk()))
        self.assertEqual(result['attachments'], [{'name': 'ander.txt', 'size': 5}])

    def test_repeat_is_idempotent_and_never_duplicates(self):
        first = self.update(body='Nieuw')
        again = self.update(body='Nieuw')
        self.assertTrue(again['already_updated']); self.assertEqual(again['id'], first['id'])
        self.assertEqual(len(self.conn.data['Drafts']), 1)
        self.assertEqual(len([c for c in self.conn.calls if c[0] == 'append']), 2)  # create + one update

    def test_refuses_non_drafts_sent_drafts_and_unrecognizable_body(self):
        calls = len(self.conn.calls)
        with self.assertRaises(ValueError): self.backend.update_draft(self.email, self.f.item, subject='x')
        with self.assertRaises(ValueError): self.backend.delete_draft(self.email, self.f.item, reason='dubbel')
        self.assertFalse(any(c[0] in ('append', 'move', 'store', 'expunge') for c in self.conn.calls[calls:]))
        self.assertIn(13, self.conn.data['INBOX'])
        manual = BytesParser(policy=policy.default).parsebytes(self.f.conn.raw()); del manual['Subject']; manual['Subject'] = 'Handmatig'
        self.conn.data['Drafts'][900] = manual.as_bytes(); manual_id = adapter.identifier(self.email, 'Drafts', 900, 7)
        with self.assertRaisesRegex(ValueError, 'cannot be delimited'): self.update(manual_id, body='x')
        self.assertIn(900, self.conn.data['Drafts'])
        self.assertEqual(self.parse(self.update(manual_id, subject='Nieuw')['id'])['Subject'], 'Nieuw')
        sent = self.backend.draft(self.email, body='Test', to='r@example.test')
        self.backend.send(self.email, sent['id'])
        with self.assertRaisesRegex(ValueError, 'submitted for sending'): self.update(sent['id'], subject='Nog eens')
        with self.assertRaisesRegex(ValueError, 'submitted for sending'): self.backend.delete_draft(self.email, sent['id'], reason='dubbel')
        self.f.smtp.sendmail.assert_called_once()

    def test_uidvalidity_change_is_reported_and_old_copy_kept(self):
        self.conn.epoch = 7
        self.conn.response = lambda name: (name, [str(self.conn.epoch).encode()] if name == 'UIDVALIDITY' else [self.conn.code.get(name)])
        original = self.conn.append
        def append(*args):
            result = original(*args); self.conn.epoch = 8
            self.conn.code['APPENDUID'] = b'8 ' + self.conn.code['APPENDUID'].split()[1]
            return result
        self.conn.append = append
        result = self.update(body='Nieuw')
        self.assertTrue(result['uidvalidity_changed']); self.assertFalse(result['old_removed']); self.assertIn('UIDVALIDITY', result['note'])
        self.assertEqual(adapter.decode(result['id'], self.email, True)['v'], 8)
        self.assertEqual(len(self.conn.data['Drafts']), 2)
        self.assertFalse(any(c[0] in ('store', 'expunge') for c in self.conn.calls))

    def test_failed_append_is_uncertain_and_never_retried_blindly(self):
        self.conn.append = Mock(side_effect=OSError('verbinding weg'))
        with self.assertRaisesRegex(ValueError, 'uncertain'): self.update(body='Nieuw')
        with self.assertRaisesRegex(ValueError, 'uncertain'): self.update(body='Nieuw')
        self.conn.append.assert_called_once(); self.assertEqual(len(self.conn.data['Drafts']), 1)

    def test_delete_draft_moves_only_from_drafts_to_trash(self):
        with self.assertRaises(ValueError): self.backend.delete_draft(self.email, self.created['id'], reason=' ')
        with self.assertRaisesRegex(ValueError, 'expect-subject'):
            self.backend.delete_draft(self.email, self.created['id'], reason='dubbel', expect_subject='Ander')
        self.assertEqual(len(self.conn.data['Drafts']), 1)
        result = self.backend.delete_draft(self.email, self.created['id'], reason='dubbel', expect_subject='Onderwerp')
        self.assertTrue(result['deleted']); self.assertFalse(result['permanent_delete'])
        self.assertEqual((len(self.conn.data['Drafts']), len(self.conn.data['Trash'])), (0, 1))
        self.assertFalse(any(c[0] == 'expunge' for c in self.conn.calls)); self.f.smtp.sendmail.assert_not_called()


class SpecTests(unittest.TestCase):
    def test_parse_update_rules(self):
        self.assertEqual(draft_edit.parse_update({'cc': ''}), {'cc': ''})
        for bad in ({}, {'attachments_remove': []}, {'body': 3}, {'to': 'a\nb'}, {'attachments_remove': ['a', 'A']}, {'x': 1}, []):
            with self.assertRaises(ValueError): draft_edit.parse_update(bad)


if __name__ == '__main__':
    unittest.main()
