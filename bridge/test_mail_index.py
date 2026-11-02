import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import mail_index as index

class IndexTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.path=self.root/'mail.sqlite'
        self.marker=patch.object(index,'marker',return_value='');self.marker.start()
        self.build=patch.object(index,'BUILD_STATUS',self.root/'build.json');self.build.start()
        self.rows=[self.row('a','a@example.test','mail','2026-10-05T08:00:00+00:00',True),
            self.row('b','a@example.test','deleted','2026-10-04T08:00:00+00:00',False),
            self.row('c','b@example.test','sent','2026-10-03T08:00:00+00:00',False)]
        self.write(self.rows)

    def tearDown(self):
        self.marker.stop();self.build.stop();self.temp.cleanup()

    def row(self,identifier,email,kind,received,unread):
        return dict(email=email,store_id='store-'+email,id=identifier,folder_id='folder-'+email,
            folder_path='Archive/Organisation',subject='Meeting finance O\'Brien',sender='Johan Example',
            sender_email='johan@example.test',received=received,unread=int(unread),kind=kind)

    def write(self,rows,complete=True):
        writer=index.SnapshotWriter(self.path)
        try:
            writer.add(rows);writer.publish(dict(complete=complete,indexed_at=index.utcnow(),invalidation='',missing=[],errors=[]))
        finally:writer.close()

    def search(self,**args):
        return index.quick_search(path=self.path,allowed={'a@example.test','b@example.test'},**args)

    def test_scope_excludes_deleted_keeps_sent_and_account_guard(self):
        result=self.search(query='finance')
        self.assertEqual(result['total'],2);self.assertTrue(result['complete'])
        self.assertEqual([m['id'] for m in result['items']],['a','c'])
        self.assertEqual(self.search(include_deleted=True)['total'],3)
        self.assertEqual(self.search(email='b@example.test')['total'],1)
        with self.assertRaises(ValueError):self.search(email='foreign@example.test')
        with self.assertRaises(ValueError):self.search(folder_id='any')

    def test_filters_combine_literal_sender_prefix_and_exclusive_utc_date(self):
        result=self.search(query="O'Brien",sender='Johan',after='2026-10-05',before='2026-10-06',unread=True)
        self.assertEqual([m['id'] for m in result['items']],['a'])
        self.assertEqual(self.search(before='2026-10-05')['total'],1)
        self.assertEqual(self.search(sender='%')['total'],0)
        self.assertEqual(self.search(query='fin')['total'],2)
        with self.assertRaises(ValueError):self.search(after='2026-10-06',before='2026-10-05')
        with self.assertRaises(ValueError):self.search(after='invalid')
        with self.assertRaises(ValueError):self.search(unread='false')

    def test_query_operators_are_literal_and_database_remains_intact(self):
        self.assertEqual(self.search(query='finance OR nonsense')['total'],0)
        self.assertEqual(self.search(query='"; DROP TABLE mail; --')['total'],0)
        self.assertEqual(self.search(query='finance')['total'],2)
        with self.assertRaises(ValueError):self.search(query='*')

    def test_pagination_exceeds_native_top_thousand_and_keeps_source(self):
        rows=[self.row(str(n),'a@example.test','mail','2026-10-05T08:00:00+00:00',False) for n in range(1200)]
        self.write(rows)
        result=self.search(offset=1000,limit=200)
        self.assertEqual(len(result['items']),200);self.assertEqual(result['total'],1200)
        self.assertIsNone(result['next_offset'])
        self.assertTrue(all(m['email']=='a@example.test' and m['folder_id']=='folder-a@example.test' for m in result['items']))

    def test_failed_snapshot_retains_good_index_without_partial_publication(self):
        with self.assertRaises(ValueError):self.write([self.row('wrong','a@example.test','mail','2026-10-05',False)],complete=False)
        self.assertEqual(self.search()['total'],2)
        self.assertEqual(list(self.root.glob('*.building-*')),[])

    def test_invalidation_marks_results_stale_and_missing_index_is_explicit(self):
        self.marker.stop();self.marker=patch.object(index,'marker',return_value='changed');self.marker.start()
        result=self.search()
        self.assertTrue(result['stale']);self.assertFalse(result['complete'])
        self.path.unlink()
        with self.assertRaisesRegex(ValueError,'being built'):self.search()

    def test_fast_queries_do_not_wait_for_native_worker_and_cannot_override_scope(self):
        import concurrent.futures
        from unittest.mock import Mock
        from server import Manager
        native=concurrent.futures.Future()
        worker=Mock();worker.submit.return_value=native
        manager=Manager(('127.0.0.1',0),backend_worker=worker)
        try:
            pending=manager.job('native','search',{'query':'body'})
            with patch('server.mail_index.combined_search',return_value={'total':7}):
                fast=manager.job('quick','quick_search',{'query':'subject'})
                self.assertEqual(fast.result(timeout=1),{'total':7})
            self.assertFalse(pending.done());worker.submit.assert_called_once()
            with self.assertRaisesRegex(Exception,'Unknown search filters'):
                manager.job('bad','quick_search',{'allowed':['foreign@example.test'],'path':'other.sqlite'})
        finally:
            manager.index_pool.shutdown();manager.imap_pool.shutdown();manager.router_pool.shutdown();manager.server_close()

if __name__=='__main__':unittest.main()
