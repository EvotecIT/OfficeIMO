"""Opt-in validator evidence and process-lifetime contracts: python3 -m unittest discover -s Build/Epub."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest

from validate_epub import ace_evidence, run_browser_audit


class AccessibilityEvidence(unittest.TestCase):
    def report(self, outcome, svg=True, version="1.4.6"):
        return {"earl:result": {"earl:outcome": outcome},
                "properties": {"hasSVGContentDocuments": svg},
                "earl:assertedBy": {"doap:release": {"doap:revision": version}}}

    def test_svg_limitation_never_overrides_a_failed_validator(self):
        for outcome, code in (("fail", 2), ("fail", 0), ("pass", 2)):
            with self.subTest(outcome=outcome, code=code):
                result = ace_evidence(self.report(outcome), code)
                self.assertEqual("failed", result["status"])
                self.assertEqual(outcome, result["outcome"])
                self.assertEqual(code, result["exitCode"])
                self.assertEqual("not-checked", result["coverage"]["svgContentDocuments"])

    def test_automated_pass_and_svg_coverage_are_independent(self):
        result = ace_evidence(self.report("pass"), 0)
        self.assertEqual("passed", result["status"])
        self.assertEqual("not-checked", result["coverage"]["svgContentDocuments"])
        self.assertEqual("ace-svg-content-not-checked", result["coverage"]["limitations"][0]["code"])
        result = ace_evidence(self.report("pass", svg=False), 0)
        self.assertEqual("passed", result["status"])
        self.assertEqual("not-applicable", result["coverage"]["svgContentDocuments"])

    def test_unknown_versions_and_missing_reports_do_not_inherit_known_coverage(self):
        for report in (self.report("pass", version="9.0.0"), {}, [], None):
            with self.subTest(report=report):
                result = ace_evidence(report, 0)
                self.assertEqual("not-established", result["coverage"]["svgContentDocuments"])
                if not isinstance(report, dict) or not report:
                    self.assertEqual("failed", result["status"])


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
