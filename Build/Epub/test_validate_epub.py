"""Opt-in process-lifetime contracts: python3 -m unittest discover -s Build/Epub."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest

from validate_epub import run_browser_audit


class BrowserAuditLifetime(unittest.TestCase):
    @unittest.skipUnless(os.name == "posix", "POSIX process-group qualification")
    def test_timeout_terminates_descendants_and_stops_evidence_writes(self):
        for detached in (False, True):
            with self.subTest(detached=detached):
                self.check_timeout(detached)

    def check_timeout(self, detached):
        with tempfile.TemporaryDirectory(prefix="epub-audit-lifetime-") as directory:
            root = Path(directory)
            child = root / "child.py"
            evidence = root / "evidence"
            child.write_text("import pathlib, sys, time\n"
                             "p = pathlib.Path(sys.argv[1])\n"
                             "while True:\n"
                             "    with p.open('a') as stream: stream.write('x')\n"
                             "    time.sleep(0.02)\n", encoding="utf-8")
            parent = root / "parent.py"
            parent.write_text("import subprocess, sys, time\n"
                              "child = subprocess.Popen([sys.executable, sys.argv[1], sys.argv[2]], start_new_session=" + repr(detached) + ")\n"
                              "print(child.pid, flush=True)\n"
                              "time.sleep(60)\n", encoding="utf-8")
            log_path = root / "process.log"
            with log_path.open("w") as log:
                with self.assertRaises(subprocess.TimeoutExpired):
                    run_browser_audit([sys.executable, str(parent), str(child), str(evidence)], log, 1)
            child_pid = int(log_path.read_text().strip())
            size = evidence.stat().st_size
            self.assertGreater(size, 0, "The descendant must actually run before the timeout")
            time.sleep(0.2)
            self.assertEqual(size, evidence.stat().st_size)
            # A reparented child may briefly remain as a zombie until the OS reaps it;
            # it has exited and cannot write evidence or consume browser resources.
            result = subprocess.run(["ps", "-o", "stat=", "-p", str(child_pid)], capture_output=True, text=True)
            self.assertTrue(not result.stdout.strip() or result.stdout.strip().startswith("Z"), result.stdout)


if __name__ == "__main__":
    unittest.main()
