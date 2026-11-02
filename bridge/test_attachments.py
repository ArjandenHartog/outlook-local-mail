import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
import mail_attachments as files
import test_manager as fixtures


class FileTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patch = patch.object(files, 'UPLOADS', self.root/'uploads'); self.patch.start()

    def tearDown(self):
        self.patch.stop(); self.temp.cleanup()

    def test_uploaded_unicode_file_and_local_path_preserve_exact_bytes(self):
        raw = b'\x00\xff real file bytes'
        upload = files.upload(io.BytesIO(raw), len(raw), 'rapport \u00e9.pdf')
        prepared = files.prepare([{'upload_id': upload['upload_id']}])
        self.assertEqual(prepared[0]['data'], raw)
        self.assertEqual(prepared[0]['mime'], 'application/pdf')
        self.assertEqual(prepared[0]['name'], 'rapport \u00e9.pdf')
        self.assertEqual(files.prepare([prepared[0]['path']])[0]['data'], raw)

    def test_paths_traversal_devices_expiry_and_incomplete_upload_rejected(self):
        for name in ('../secret','C:\\secret','NUL.txt','file.','evil\r\n.txt'):
            with self.assertRaises(ValueError): files.upload(io.BytesIO(b''),0,name)
        for value in ('relative.pdf', {'upload_id':'../secret'}, {'upload_id':'0'*32}, {'upload_id':'0'*32,'path':'evil'}):
            with self.assertRaises(ValueError): files.prepare([value])
        with self.assertRaises(ValueError): files.upload(io.BytesIO(b'a'),2,'partial.txt')
        upload = files.upload(io.BytesIO(b'a'),1,'expired.txt')
        with patch.object(files,'TTL_SECONDS',-1):
            with self.assertRaises(ValueError): files.prepare([{'upload_id':upload['upload_id']}])

    def test_total_size_and_count_validated_before_mail_mutation(self):
        a=self.root/'a.txt';a.write_bytes(b'abc')
        with patch.object(files,'MAX_BYTES',5):
            with self.assertRaises(ValueError): files.prepare([str(a),str(a)])
        with self.assertRaises(ValueError): files.prepare([str(a)]*21)
        with self.assertRaises(ValueError): files.upload(io.BytesIO(),files.MAX_BYTES+1,'large.pdf')


class NativeAttachmentsTests(unittest.TestCase):
    def setUp(self):
        self.fixture=fixtures.BackendTests();self.fixture.setUp()

    def tearDown(self): self.fixture.tearDown()

    def test_native_attaches_copy_to_correct_mailbox_and_returns_receipt(self):
        path=Path(self.fixture.temp.name)/'report.pdf';path.write_bytes(b'%PDF-fixture')
        created=[]
        message_class=fixtures.Message
        def create(*args):
            message=message_class(*args);message.Attachments=Mock();created.append(message);return message
        with patch('test_manager.Message',side_effect=create):
            result=self.fixture.backend.draft('a@example.test',body='Ha',signature='none',attachments=[str(path)])
        created[0].Attachments.Add.assert_called_once_with(str(path.resolve()),1)
        self.assertFalse(result['sent']);self.assertEqual(result['email'],'a@example.test')
        self.assertEqual(result['attachments'],[{'name':'report.pdf','size':len(b'%PDF-fixture')}])

    def test_missing_file_does_not_create_native_draft(self):
        with self.assertRaises(ValueError):
            self.fixture.backend.draft('a@example.test',attachments=[str(Path(self.fixture.temp.name)/'missing.pdf')])
        self.assertEqual(len(self.fixture.drafts.Items),0)

    def test_native_table_reads_pages_without_opening_individual_messages(self):
        table=Mock();table.GetRowCount.return_value=3
        received=self.fixture.item.ReceivedTime
        table.GetArray.side_effect=[[(1,)], [('id2','Hello','Johan','EX-address','johan@example.test',received,True)]]
        self.fixture.inbox.GetTable=Mock(return_value=table)
        result=self.fixture.backend.messages('a@example.test',offset=1,limit=1)
        self.assertEqual(result['next_offset'],2);self.assertEqual(result['total'],3)
        self.assertEqual(result['items'][0]['sender_email'],'johan@example.test')
        self.assertEqual(result['source'],'Outlook metadata table')
        table.GetArray.assert_any_call(1)


if __name__=='__main__': unittest.main()
