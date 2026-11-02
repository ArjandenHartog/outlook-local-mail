import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock,patch
from email.message import EmailMessage
from email import policy
import imap_backend as adapter
import mail_index

class FakeImap:
    def __init__(self):
        self.folder='INBOX';self.readonly=True;self.calls=[];self.code={};self.next_uid=100
        self.data={'INBOX':{13:self.raw()},'Archive':{},'Drafts':{},'Sent':{},'Trash':{}}
        self.flags={('INBOX',13):[]}

    def raw(self):
        m=EmailMessage();m['From']='Johan <johan@example.test>';m['To']='me@example.com'
        m['Subject']='Question';m['Message-ID']='<original@example.test>';m.set_content('Hello there')
        return m.as_bytes()

    def noop(self):return 'OK',[]
    def logout(self):return 'BYE',[]
    def list(self):return 'OK',[b'(\\HasNoChildren) "." INBOX',b'(\\HasNoChildren) "." Archive',b'(\\HasNoChildren \\Drafts) "." Drafts',b'(\\HasNoChildren \\Sent) "." Sent',b'(\\HasNoChildren \\Trash) "." Trash']
    def select(self,name,readonly=True):
        self.folder=name.strip('"');self.readonly=readonly;self.calls.append(('select',name,readonly))
        return 'OK',[str(len(self.data[self.folder])).encode()]
    def response(self,name):return name,[b'7'] if name=='UIDVALIDITY' else [self.code.get(name)]
    def capability(self):return 'OK',[b'IMAP4rev1 UIDPLUS MOVE']
    def status(self,name,fields):return 'OK',[b'INBOX (MESSAGES 1 UNSEEN 1)']
    def append(self,name,flags,date,raw):
        self.next_uid+=1;folder=name.strip('"');self.data[folder][self.next_uid]=raw
        self.flags[(folder,self.next_uid)]=[b'\\Seen',b'\\Draft'] if folder=='Drafts' else [b'\\Seen']
        self.code['APPENDUID']=b'7 '+str(self.next_uid).encode();self.calls.append(('append',folder))
        return 'OK',[b'append done']
    def uid(self,command,*args):
        self.calls.append((command,)+args)
        if command=='search':return 'OK',[b' '.join(str(n).encode() for n in self.data[self.folder])]
        if command=='fetch':
            values=args[0].decode() if isinstance(args[0],bytes) else str(args[0]);rows=[]
            for uid in map(int,values.split(',')):
                if uid not in self.data[self.folder]:continue
                flags=b' '.join(self.flags.get((self.folder,uid),[]))
                attrs=b'1 (UID '+str(uid).encode()+b' FLAGS ('+flags+b') INTERNALDATE "05-Oct-2026 10:00:00 +0200" BODY[] {1}'
                rows.extend([(attrs,self.data[self.folder][uid]),b')'])
            return 'OK',rows
        if command=='move':
            if self.readonly:raise AssertionError('Move on readonly connection')
            uid=int(args[0]);target=args[1].strip('"');self.next_uid+=1
            self.data[target][self.next_uid]=self.data[self.folder].pop(uid)
            self.flags[(target,self.next_uid)]=self.flags.pop((self.folder,uid),[])
            self.code['COPYUID']=b'7 '+str(uid).encode()+b' '+str(self.next_uid).encode()
            return 'OK',[]
        if command=='store':
            if self.readonly:raise AssertionError('Store on readonly connection')
            return 'OK',[]
        if command=='expunge':self.data[self.folder].pop(int(args[0]),None);return 'OK',[]
        raise AssertionError(command)

class ImapTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);(self.root/'runtime').mkdir()
        self.patch=patch.object(adapter,'ROOT',self.root);self.patch.start()
        self.marker=patch.object(adapter,'IMAP_MARKER',self.root/'runtime/imap-invalidated.json');self.marker.start()
        logo=base64.b64encode(b'fixture image').decode()
        (self.root/'runtime/signature.html').write_text('<p>Jane Doe {{email}}</p><img src="data:image/png;base64,'+logo+'">',encoding='utf-8')
        self.config={'email':'me@example.com','username':'me@example.com','aliases':['alias@example.com'],
            'host':'mail.example.test','imap_port':993,'smtp_port':465,'display_name':'Jane Doe','credential_file':'fixture.dpapi','signature_file':'signature.html','signature_name':'Default'}
        self.conn=FakeImap();self.smtp=Mock();self.smtp.sendmail.return_value={}
        self.smtp_context=Mock();self.smtp_context.__enter__=Mock(return_value=self.smtp);self.smtp_context.__exit__=Mock(return_value=False)
        self.factory=Mock(return_value=self.smtp_context)
        self.backend=adapter.ImapBackend([self.config],connector=lambda config:self.conn,smtp_factory=self.factory)
        self.backend.secret=Mock(return_value='test credential')
        self.item=adapter.identifier(self.config['email'],'INBOX',13,7)

    def tearDown(self):self.marker.stop();self.patch.stop();self.temp.cleanup()

    def test_unicode_folder_roundtrip_and_foreign_message_epoch_rejected(self):
        self.assertEqual(adapter.utf7_decode(adapter.utf7_encode('Priv\u00e9 & Klanten')),'Priv\u00e9 & Klanten')
        self.assertEqual(adapter.parse_list(b'(\\HasNoChildren) "." Priv&AOk-')['name'],'Priv\u00e9')
        with self.assertRaises(ValueError):adapter.decode(self.item,'someone@example.test',message=True)
        with self.assertRaises(ValueError):self.backend.read(self.config['email'],adapter.identifier(self.config['email'],'INBOX',13,8))

    def test_read_and_index_headers_are_peek_and_preserve_unread_utc(self):
        result=self.backend.read(self.config['email'],self.item)
        self.assertTrue(result['unread']);self.assertEqual(result['received'],'2026-10-05T08:00:00+00:00')
        self.assertIn('Hello there',result['body'])
        self.assertTrue(all(call[2] is True for call in self.conn.calls if call[0]=='select'))
        self.assertFalse(any(call[0] in ('store','move','expunge') for call in self.conn.calls))
        self.assertTrue(any('BODY.PEEK[]' in call[-1] for call in self.conn.calls if call[0]=='fetch'))
        rows=self.backend.metadata(self.conn,self.config['email'],{'wire':'INBOX','kind':'mail','name':'INBOX'})
        self.assertEqual(rows[0]['sender_email'],'johan@example.test')
        self.assertTrue(any('BODY.PEEK[HEADER.FIELDS' in call[-1] for call in self.conn.calls if call[0]=='fetch'))

    def test_alias_draft_cid_signature_and_reply_headers(self):
        result=self.backend.draft(self.config['email'],body='Ha Johan',reply_id=self.item,from_address='alias@example.com')
        self.assertFalse(result['sent']);self.assertEqual(result['sender_email'],'alias@example.com')
        draft=self.backend.read(self.config['email'],result['id']);self.assertTrue(draft['is_draft'])
        raw=self.conn.data['Drafts'][adapter.decode(result['id'],self.config['email'],True)['u']]
        message=adapter.BytesParser(policy=policy.default).parsebytes(raw)
        self.assertIn('alias@example.com',str(message['From']));self.assertEqual(message['In-Reply-To'],'<original@example.test>')
        self.assertEqual(len(draft['attachments']),1)
        self.assertIn('cid:signature-',message.get_body(preferencelist=('html',)).get_content())
        self.smtp.sendmail.assert_not_called()

    def test_unauthorized_alias_or_header_injection_rejected_before_append(self):
        before=len(self.conn.calls)
        with self.assertRaises(ValueError):self.backend.draft(self.config['email'],from_address='admin@example.com')
        with self.assertRaises(ValueError):self.backend.draft(self.config['email'],subject='Hello\r\nBcc: attacker@example.test')
        self.assertFalse(any(call[0]=='append' for call in self.conn.calls[before:]))

    def test_files_coexist_with_inline_logo_and_alias_without_smtp(self):
        path=self.root/'rapport \u00e9.pdf';content=b'%PDF-1.7\x00\xff';path.write_bytes(content)
        result=self.backend.draft(self.config['email'],body='Ha',from_address='alias@example.com',attachments=[str(path)])
        uid=adapter.decode(result['id'],self.config['email'],True)['u']
        message=adapter.BytesParser(policy=policy.default).parsebytes(self.conn.data['Drafts'][uid])
        attached=list(message.iter_attachments())
        self.assertEqual(len(attached),1);self.assertEqual(attached[0].get_filename(),path.name)
        self.assertEqual(attached[0].get_payload(decode=True),content)
        self.assertEqual(attached[0].get_content_type(),'application/pdf')
        self.assertTrue(any(p.get_content_disposition()=='inline' for p in message.walk()))
        self.assertIn('cid:signature-',message.get_body(preferencelist=('html',)).get_content())
        self.smtp.sendmail.assert_not_called()

    def test_missing_file_does_not_append_imap_draft(self):
        with self.assertRaises(ValueError):self.backend.draft(self.config['email'],attachments=[str(self.root/'missing.pdf')])
        self.assertFalse(any(c[0]=='append' for c in self.conn.calls))

    def test_unchanged_index_skips_fetch_but_modseq_epoch_and_force_refresh(self):
        self.conn.capability=lambda:('OK',[b'IMAP4rev1 CONDSTORE'])
        fingerprint=[1,7,14,1,50]
        def status(name,fields):
            values=' '.join(f'{key} {value}' for key,value in zip(('MESSAGES','UIDVALIDITY','UIDNEXT','UNSEEN','HIGHESTMODSEQ'),fingerprint))
            return 'OK',[('INBOX ('+values+')').encode()]
        self.conn.status=status;folder={'wire':'INBOX','kind':'mail','name':'INBOX'}
        first=self.backend.indexed_metadata(self.conn,self.config['email'],folder)
        fetches=lambda:len([c for c in self.conn.calls if c[0]=='fetch'])
        before=fetches();self.assertEqual(first,self.backend.indexed_metadata(self.conn,self.config['email'],folder))
        self.assertEqual(before,fetches())
        # Seen changes can swap between messages without changing unread count.
        fingerprint[-1]+=1;self.conn.flags[('INBOX',13)]=[b'\\Seen']
        changed=self.backend.indexed_metadata(self.conn,self.config['email'],folder)
        self.assertFalse(changed[0]['unread']);self.assertGreater(fetches(),before)
        before=fetches();fingerprint[1]+=1
        self.backend.indexed_metadata(self.conn,self.config['email'],folder);self.assertGreater(fetches(),before)
        before=fetches();self.backend.indexed_metadata(self.conn,self.config['email'],folder,force=True)
        self.assertGreater(fetches(),before)

    def test_zero_modseq_fetches_each_time_and_racing_change_rejects_cache(self):
        folder={'wire':'INBOX','kind':'mail','name':'INBOX'}
        self.conn.capability=lambda:('OK',[b'CONDSTORE'])
        self.conn.status=lambda *args:('OK',[b'INBOX (MESSAGES 1 UIDVALIDITY 7 UIDNEXT 14 UNSEEN 1 HIGHESTMODSEQ 0)'])
        self.backend.indexed_metadata(self.conn,self.config['email'],folder)
        before=len(self.conn.calls);self.backend.indexed_metadata(self.conn,self.config['email'],folder)
        self.assertTrue(any(c[0]=='fetch' for c in self.conn.calls[before:]))
        with patch.object(self.backend,'folder_fingerprint',side_effect=[(1,7,14,1,1),(1,7,14,0,2)]):
            with self.assertRaises(ValueError):self.backend.indexed_metadata(self.conn,self.config['email'],folder)
        self.assertFalse(self.backend.metadata_cache)

    def test_send_logs_in_primary_uses_alias_envelope_and_is_durable(self):
        draft=self.backend.draft(self.config['email'],body='Test',to='recipient@example.test',from_address='alias@example.com')
        result=self.backend.send(self.config['email'],draft['id'])
        self.assertTrue(result['submitted']);self.assertTrue(result['saved_sent']);self.assertTrue(result['draft_removed'])
        self.smtp.login.assert_called_once_with('me@example.com','test credential')
        self.assertEqual(self.smtp.sendmail.call_args.args[:2],('alias@example.com',['recipient@example.test']))
        self.assertEqual(self.backend.send(self.config['email'],draft['id'])['state'],'submitted');self.smtp.sendmail.assert_called_once()
        journal=(self.root/'runtime/imap-send-journal.json').read_text('utf-8');self.assertNotIn('test credential',journal)

    def test_uncertain_send_does_not_retry_and_non_draft_cannot_send(self):
        with self.assertRaises(ValueError):self.backend.send(self.config['email'],self.item)
        draft=self.backend.draft(self.config['email'],body='Test',to='recipient@example.test')
        self.smtp.sendmail.side_effect=OSError('connection lost after DATA')
        for _ in range(2):
            with self.assertRaises(ValueError):self.backend.send(self.config['email'],draft['id'])
        self.smtp.sendmail.assert_called_once()

    def test_archive_move_uses_uid_move_preserves_unread_and_blocks_trash(self):
        with self.assertRaises(ValueError):self.backend.move(self.config['email'],self.item,adapter.identifier(self.config['email'],'Trash'))
        result=self.backend.move(self.config['email'],self.item,adapter.identifier(self.config['email'],'Archive'),adapter.identifier(self.config['email'],'INBOX'))
        self.assertTrue(self.backend.read(self.config['email'],result['new_id'])['unread'])
        repeat=self.backend.move(self.config['email'],self.item,adapter.identifier(self.config['email'],'Archive'))
        self.assertTrue(repeat['already_moved']);self.assertEqual(len([call for call in self.conn.calls if call[0]=='move']),1)

    def test_trash_moves_inbox_to_own_trash_preserves_unread_and_is_idempotent(self):
        inbox=adapter.identifier(self.config['email'],'INBOX')
        with self.assertRaises(ValueError):self.backend.trash(self.config['email'],self.item,inbox,' ')
        with self.assertRaises(ValueError):self.backend.trash(self.config['email'],self.item,adapter.identifier(self.config['email'],'Archive'),'reclame')
        result=self.backend.trash(self.config['email'],self.item,inbox,'reclame')
        self.assertFalse(result['permanent_delete']);self.assertEqual(result['reason'],'reclame')
        self.assertEqual(adapter.decode(result['new_id'],self.config['email'],True)['f'],'Trash')
        self.assertTrue(self.backend.read(self.config['email'],result['new_id'])['unread'])
        self.assertFalse(any(call[0] in ('expunge','store') for call in self.conn.calls))
        self.assertTrue(self.backend.trash(self.config['email'],self.item,inbox,'reclame')['already_moved'])
        self.assertEqual(len([call for call in self.conn.calls if call[0]=='move']),1)

    def test_trash_rejects_message_outside_inbox(self):
        moved=self.backend.move(self.config['email'],self.item,adapter.identifier(self.config['email'],'Archive'))
        with self.assertRaises(ValueError):
            self.backend.trash(self.config['email'],moved['new_id'],adapter.identifier(self.config['email'],'INBOX'),'reclame')
        self.assertEqual(len([call for call in self.conn.calls if call[0]=='move']),1)

    def test_combined_search_global_page_account_scopes_and_missing_provider(self):
        a=self.root/'a.sqlite';b=self.root/'b.sqlite'
        for path,email,date in ((a,'a@example.test','2026-10-04T08:00:00+00:00'),(b,'b@example.test','2026-10-05T08:00:00+00:00')):
            writer=mail_index.SnapshotWriter(path)
            try:
                writer.add([dict(email=email,store_id='store',id=email,folder_id='folder',folder_path='Archive',subject='Meeting',sender='Sender',sender_email='sender@example.test',received=date,unread=False,kind='mail')])
                writer.publish({'complete':True,'indexed_at':mail_index.utcnow(),'invalidation':''})
            finally:writer.close()
        with patch('mail_index.sources',return_value=[(a,{'a@example.test'}),(b,{'b@example.test'})]),patch('mail_index.marker',return_value=''):
            result=mail_index.combined_search(query='Meeting',limit=1)
            self.assertEqual(result['total'],2);self.assertEqual(result['items'][0]['email'],'b@example.test');self.assertEqual(result['next_offset'],1)
            self.assertEqual(mail_index.combined_search(query='Meeting',limit=1,offset=1)['items'][0]['email'],'a@example.test')
            b.unlink();result=mail_index.combined_search(query='Meeting')
            self.assertFalse(result['complete']);self.assertEqual(result['missing'],['b@example.test'])

if __name__=='__main__':unittest.main()
