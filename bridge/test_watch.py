"""Tests for skill/scripts/watch.py with a fake service; nothing is read from or written to a mailbox."""
import datetime as dt
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'skill/scripts'))
import watch  # noqa: E402


def days_ago(days):
    return (dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=days)).isoformat()


class FakeService:
    def __init__(self):
        self.calls = []
        self.drafts = [
            {'id': 'd1', 'subject': 'Offer', 'received': days_ago(0.1)},
            {'id': 'd2', 'subject': 'RE: Offer', 'received': days_ago(0.5)},
            {'id': 'd3', 'subject': 'Different', 'received': days_ago(9)},
            {'id': 'd4', 'subject': 'Booking 12345', 'received': days_ago(0.2)},
            {'id': 'd5', 'subject': 'Booking 12345', 'received': days_ago(0.3)},
        ]
        self.bodies = {'d1': 'Hello, see attached.', 'd2': 'Hello, see attached.', 'd3': 'x', 'd4': 'b', 'd5': 'b'}
        self.sent = [{'id': 's%d' % i, 'subject': 'News', 'received': days_ago(0.01 * i)} for i in range(1, 5)]

    def __call__(self, action, args=None, polls=60):
        args = args or {}
        self.calls.append(action)
        if action == 'status':
            return {'accounts': [{'email': 'a@example.test', 'connected': True}, {'email': 'b@example.test', 'connected': False}]}
        if action == 'folders':
            return {'drafts_id': 'F-drafts', 'sent_id': 'F-sent'}
        if action == 'messages':
            items = self.drafts if args['folder_id'] == 'F-drafts' else self.sent
            return {'items': items, 'total': len(items), 'next_offset': None}
        if action == 'read':
            return {'to': 'x@example.test', 'body': self.bodies[args['item_id']], 'attachments': []}
        raise AssertionError(action)


class WatchTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.service = FakeService()
        for patcher in (patch.object(watch, 'call', self.service), patch.object(watch, 'RUNTIME', root),
                        patch.object(watch, 'STATE', root / 'state.json'), patch.object(watch, 'CONFIG', root / 'config.json')):
            patcher.start(); self.addCleanup(patcher.stop)
        (root / 'config.json').write_text(json.dumps({'ignore_drafts': [{'account': 'A@example.test', 'subject_regex': 'booking 12345'}]}))

    def tearDown(self):
        self.temp.cleanup()

    def test_draft_audit_groups_duplicates_flags_stale_and_honours_ignore_list(self):
        report = watch.draft_audit(NS(account=None, stale_days=7))
        self.assertEqual(report['missing_accounts'], ['b@example.test'])
        self.assertEqual(report['ignored'], 2)
        self.assertEqual(len(report['duplicate_groups']), 1)
        group = report['duplicate_groups'][0]
        self.assertEqual((group['keep_newest'], group['suggested_delete']), ('d1', ['d2']))
        self.assertEqual([d['id'] for d in report['stale_drafts']], ['d3'])
        self.assertFalse(report['sends_or_deletes_anything'])
        self.assertTrue(set(self.service.calls) <= {'status', 'folders', 'messages', 'read'})
        self.assertNotIn('body', json.dumps(report))

    def test_different_text_or_recipients_is_not_suggested_for_deletion(self):
        self.service.bodies['d2'] = 'Completely different message text'
        group = watch.draft_audit(NS(account=None, stale_days=7))['duplicate_groups'][0]
        self.assertEqual(group['suggested_delete'], [])

    def verify(self, outcomes):
        def fake(args, wait=0, interval=20):
            return outcomes[args['item_id']]
        return patch.object(watch.mail, 'verify_sent', side_effect=fake)

    def test_bounce_watch_alerts_once_and_reports_ambiguous_ndr_once(self):
        outcomes = {'s1': {'bounce_detected': True, 'failed_recipients': ['x@example.test'], 'explanation': 'failed'},
                    's2': {'possible_bounces': [{'received': 'r', 'subject': 'Undeliverable: News', 'folder_path': 'Deleted'}]},
                    's3': {'delivery_verified': True},
                    's4': {'possible_bounces': [{'received': 'r', 'subject': 'Undeliverable: News', 'folder_path': 'Deleted'}]}}
        args = NS(account=None, days=3, max=40, recheck_hours=3)
        with self.verify(outcomes):
            first = watch.bounce_watch(args)
            second = watch.bounce_watch(args)
        self.assertEqual([b['failed_recipients'] for b in first['new_bounces']], [['x@example.test']])
        self.assertEqual(len(first['ambiguous_ndrs']), 1)
        self.assertEqual((first['verified'], first['delivered']), (4, 1))
        self.assertEqual(second['new_bounces'], []); self.assertEqual(second['ambiguous_ndrs'], [])
        self.assertEqual(second['verified'], 0)  # settled, or pending and rechecked too recently
        self.assertNotIn('s1', (Path(watch.STATE).read_text()))

    def test_cap_checks_unseen_messages_first_next_run(self):
        outcomes = {i: {} for i in ('s1', 's2', 's3', 's4')}
        args = NS(account=None, days=3, max=2, recheck_hours=0)
        checked = []
        def fake(call_args, wait=0, interval=20):
            checked.append(call_args['item_id']); return outcomes[call_args['item_id']]
        with patch.object(watch.mail, 'verify_sent', side_effect=fake):
            first = watch.bounce_watch(args)
            second = watch.bounce_watch(args)
        self.assertEqual(first['skipped_over_limit'], 2)
        self.assertEqual(len(set(checked[:2]) & set(checked[2:])), 0)  # second run picked the two unseen ones


if __name__ == '__main__':
    unittest.main()
