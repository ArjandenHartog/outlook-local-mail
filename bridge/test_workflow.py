import datetime as dt
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import Mock, patch
from outlook_backend import OutlookBackend, MailError
from test_manager import Collection, Message, fake_folder

class Table:
    def __init__(self,rows):
        self.rows=rows;self.Columns=NS(RemoveAll=lambda:None,Add=lambda col:None)
    def GetRowCount(self):return len(self.rows)
    def Sort(self,*args):pass
    def GetArray(self,count):return self.rows[:count]

class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();root=Path(self.temp.name)
        self.accounts=root/'accounts.json';self.accounts.write_text(json.dumps([{'email':'a@example.test'}]))
        self.folders={}
        for name in ('root','inbox','archive','deleted','junk','outbox','drafts','sent','virtual'):
            f=fake_folder('store-a',name,name);f.DefaultItemType=0
            f.PropertyAccessor=NS(GetProperty=lambda tag,n=name:2 if n=='virtual' else 1)
            f.GetTable=Mock(return_value=Table([]));self.folders[name]=f
        self.root=self.folders['root'];self.root.Folders.extend(f for k,f in self.folders.items() if k!='root')
        for f in self.root.Folders:f.Parent=self.root
        self.store=NS(StoreID='store-a',GetRootFolder=lambda:self.root,
            GetDefaultFolder=lambda n:self.folders[{3:'deleted',4:'outbox',5:'sent',6:'inbox',16:'drafts',23:'junk'}[n]])
        self.account=NS(SmtpAddress='a@example.test',DeliveryStore=self.store)
        self.item=Message('old',self.folders['inbox']);self.items={'old':self.item}
        self.session=NS(Accounts=Collection([self.account]),GetFolderFromID=lambda i,s:self.folders[i],
            GetItemFromID=lambda i,s:self.items[i])
        self.backend=OutlookBackend(self.accounts,NS(Session=self.session),journal_path=root/'send-journal.json')

    def tearDown(self):self.temp.cleanup()

    def test_search_finds_archive_and_sent_but_excludes_deleted_junk_and_virtual(self):
        date=dt.datetime(2026,10,5,8,tzinfo=dt.timezone.utc)
        row=lambda i:(i,'Invoice','Seller','seller@example.test',date,True)
        self.folders['archive'].GetTable.return_value=Table([row('a')])
        self.folders['sent'].GetTable.return_value=Table([row('s')])
        for n in ('deleted','junk','virtual'):self.folders[n].GetTable.return_value=Table([row(n)])
        result=self.backend.search("O'Brien",email='a@example.test')
        self.assertEqual(result['total'],2);self.assertTrue(result['complete'])
        self.assertEqual({m['folder_path'] for m in result['items']},{'archive','sent'})
        self.assertTrue(all(m['received']=='2026-10-05T08:00:00+00:00' for m in result['items']))
        self.assertIn("O''Brien",self.folders['archive'].GetTable.call_args.args[0])
        for n in ('deleted','junk','virtual'):self.folders[n].GetTable.assert_not_called()

    def test_search_failure_is_partial_and_cursor_cannot_switch_mailbox_or_query(self):
        self.folders['archive'].GetTable.side_effect=RuntimeError('unavailable')
        result=self.backend.search('invoice')
        self.assertTrue(result['scan_complete']);self.assertFalse(result['complete']);self.assertEqual(len(result['errors']),1)
        with self.assertRaises(MailError):self.backend.search('other',cursor=result['cursor'])
        with self.assertRaises(MailError):self.backend.search('invoice',email='unknown@example.test')

    def test_move_rejects_foreign_destination_and_changed_source_without_mutation(self):
        foreign=fake_folder('store-b','foreign','foreign');self.folders['foreign']=foreign
        self.item.Move=Mock()
        with self.assertRaises(MailError):self.backend.move('a@example.test','old','foreign')
        with self.assertRaises(MailError):self.backend.move('a@example.test','old','archive',source_folder_id='changed')
        self.item.Move.assert_not_called()

    def test_move_preserves_unread_new_id_and_is_durable_without_old_id(self):
        def native_move(target):
            self.item.Parent=target;self.item.UnRead=False;self.item.EntryID='new';self.items['new']=self.item;del self.items['old'];return self.item
        move=Mock(side_effect=native_move);self.item.Move=move
        result=self.backend.move('a@example.test','old','archive',source_folder_id='inbox')
        self.assertTrue(self.item.UnRead);self.assertEqual(result['new_id'],'new');self.assertEqual(self.item.saves,1)
        fresh=OutlookBackend(self.accounts,NS(Session=self.session),journal_path=self.backend.journal_path)
        retry=fresh.move('a@example.test','old','archive')
        self.assertTrue(retry['already_moved']);self.assertEqual(move.call_count,1)

    def test_uncertain_move_never_retries_automatically(self):
        self.item.Move=Mock(side_effect=RuntimeError('disconnect after move'))
        with self.assertRaises(MailError):self.backend.move('a@example.test','old','archive')
        with self.assertRaises(MailError):self.backend.move('a@example.test','old','archive')
        self.assertEqual(self.item.Move.call_count,1)

    def test_move_is_archiving_and_does_not_delete_or_move_pending_mail(self):
        self.item.Move=Mock()
        for name in ('deleted','junk','outbox','drafts'):
            with self.assertRaises(MailError):self.backend.move('a@example.test','old',name)
        self.item.Parent=self.folders['outbox']
        with self.assertRaises(MailError):self.backend.move('a@example.test','old','archive')
        self.item.Move.assert_not_called()

    def test_trash_only_moves_inbox_to_own_deleted_and_preserves_unread(self):
        def native_move(target):
            self.item.Parent=target;self.item.UnRead=False;self.item.EntryID='new';self.items['new']=self.item;del self.items['old'];return self.item
        self.item.Move=Mock(side_effect=native_move)
        result=self.backend.trash('a@example.test','old','inbox','Duidelijke reclame')
        self.assertEqual(result['target_folder_id'],'deleted')
        self.assertFalse(result['permanent_delete']);self.assertTrue(self.item.UnRead)
        again=self.backend.trash('a@example.test','old','inbox','Duidelijke reclame')
        self.assertTrue(again['already_moved']);self.assertEqual(self.item.Move.call_count,1)

    def test_trash_rejects_archive_changed_source_and_empty_reason(self):
        self.item.Move=Mock()
        with self.assertRaises(MailError):self.backend.trash('a@example.test','old','archive','Reclame')
        with self.assertRaises(MailError):self.backend.trash('a@example.test','old','inbox','')
        self.item.Parent=self.folders['archive']
        with self.assertRaises(MailError):self.backend.trash('a@example.test','old','inbox','Reclame')
        self.item.Move.assert_not_called()

    def test_suggestions_never_use_deleted_history_or_another_mailbox(self):
        self.item.SenderEmailAddress='sender@example.test';self.item.SenderName='Sender'
        analysis={'checked_at':'2026-10-05','mailboxes':[{'email':'a@example.test',
            'folders':[{'id':'archive','path':'archive','kind':'archive'},{'id':'deleted','path':'deleted','kind':'deleted'}],
            'senders':{'sender@example.test':{'folders':{'archive':5,'deleted':300},'recent':[{'name':'Sender'}]}}}]}
        runtime=Path(self.temp.name)/'runtime';runtime.mkdir();(runtime/'filing-analysis.json').write_text(json.dumps(analysis))
        with patch('workflow.ROOT',Path(self.temp.name)):
            result=self.backend.suggest_folder('a@example.test','old')
            self.assertEqual(len(result['suggestions']),1)
            self.assertEqual(result['suggestions'][0]['total_examples'],5)
            self.assertEqual(result['suggestions'][0]['confidence'],'strong')
            with self.assertRaises(MailError):self.backend.suggest_folder('unknown@example.test','old')

    def test_read_shows_native_smtp_sender_without_confusing_receiver_account(self):
        self.item.SenderEmailAddress='/O=EXCHANGE/CN=OTHER-PERSON'
        self.item.SendUsingAccount=self.account
        self.item.PropertyAccessor=NS(GetProperty=lambda tag:'sender@company.test' if '0x5D01' in tag else self.item.ReceivedTime)
        self.assertEqual(self.backend.read('a@example.test','old')['sender_email'],'sender@company.test')

if __name__=='__main__':unittest.main()
