import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import supervisor


class FilingScheduleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / 'runtime').mkdir()
        for name, value in (('ROOT', self.root), ('RUNTIME', self.root / 'runtime')):
            patcher = patch.object(supervisor, name, value); patcher.start(); self.addCleanup(patcher.stop)
        self.popen = patch.object(supervisor.subprocess, 'Popen'); self.launch = self.popen.start(); self.addCleanup(self.popen.stop)

    def tearDown(self):
        self.temp.cleanup()

    def test_nothing_runs_without_configured_outlook_accounts(self):
        supervisor.schedule_filing()
        self.launch.assert_not_called()

    def test_first_start_builds_analysis_once_then_waits_before_retrying(self):
        (self.root / 'accounts.json').write_text('[]')
        supervisor.schedule_filing(); supervisor.schedule_filing()
        self.launch.assert_called_once()
        self.assertTrue(str(self.launch.call_args.args[0][1]).endswith('analyze_filing.py'))
        self.assertEqual(self.launch.call_args.kwargs['cwd'], self.root)

    def test_fresh_analysis_is_kept_and_old_one_is_refreshed(self):
        (self.root / 'accounts.json').write_text('[]')
        analysis = self.root / 'runtime/filing-analysis.json'; analysis.write_text('{}')
        supervisor.schedule_filing(); self.launch.assert_not_called()
        old = time.time() - 8 * 86400; os.utime(analysis, (old, old))
        supervisor.schedule_filing(); self.launch.assert_called_once()


if __name__ == '__main__':
    unittest.main()
