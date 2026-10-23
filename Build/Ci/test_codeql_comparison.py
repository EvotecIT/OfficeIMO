"""Coverage evidence must detect lost source and changed alerts, not just counts."""
import json
from pathlib import Path
import tempfile
import unittest

from compare_codeql import compare


class CodeQLComparisonTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.before_files = self.root / "before.csv"
        self.after_files = self.root / "after.csv"
        self.before_sarif = self.root / "before.sarif"
        self.after_sarif = self.root / "after.sarif"
        self.expected_rules = self.root / "expected-rules.json"
        self.expected_rules.write_text('["cs/path-injection"]', encoding="utf-8")
        for path in [self.before_files, self.after_files]:
            path.write_text('file,lines\nengine/Document.cs,120\n', encoding="utf-8")
        self.sarif = {"runs": [{"tool": {"driver": {"name": "CodeQL", "rules": [{"id": "cs/path-injection"}]}},
                              "invocations": [{"executionSuccessful": True}],
                              "results": [{"ruleId": "cs/path-injection", "message": {"text": "Unsafe path"},
                                           "locations": [{"physicalLocation": {"artifactLocation": {"uri": "engine/Document.cs"},
                                                                               "region": {"startLine": 12}}}]}]}]}
        for path in [self.before_sarif, self.after_sarif]:
            path.write_text(json.dumps(self.sarif), encoding="utf-8")

    def result(self):
        return compare(self.before_files, self.after_files, self.before_sarif, self.after_sarif, self.expected_rules)

    def test_same_extraction_and_alerts_are_equal(self):
        self.assertTrue(self.result()["equal"])

    def test_full_suite_includes_metric_rules_and_diagnostic_notifications(self):
        driver = self.sarif["runs"][0]["tool"]["driver"]
        driver["rules"].append({"id": "cs/summary/lines-of-code"})
        driver["notifications"] = [{"id": "cs/diagnostics/extraction-errors"}]
        self.expected_rules.write_text('["cs/path-injection", "cs/summary/lines-of-code", "cs/diagnostics/extraction-errors"]', encoding="utf-8")
        for path in [self.before_sarif, self.after_sarif]:
            path.write_text(json.dumps(self.sarif), encoding="utf-8")
        self.assertTrue(self.result()["equal"])

    def test_missing_diagnostic_query_cannot_pass(self):
        self.expected_rules.write_text('["cs/path-injection", "cs/diagnostics/extraction-errors"]', encoding="utf-8")
        self.sarif["runs"][0]["tool"]["driver"]["notifications"] = [{"id": "cs/diagnostics/extraction-errors"}]
        self.before_sarif.write_text(json.dumps(self.sarif), encoding="utf-8")
        self.assertFalse(self.result()["equal"])

    def test_same_count_source_substitution_is_rejected(self):
        self.after_files.write_text('file,lines\nengine/Other.cs,120\n', encoding="utf-8")
        result = self.result()
        self.assertFalse(result["equal"])
        self.assertEqual([["engine/Document.cs", "120"]], [list(row) for row in result["missing_files"]])

    def test_changed_alert_location_is_rejected(self):
        self.sarif["runs"][0]["results"][0]["locations"][0]["physicalLocation"]["region"]["startLine"] = 13
        self.after_sarif.write_text(json.dumps(self.sarif), encoding="utf-8")
        self.assertFalse(self.result()["equal"])

    def test_empty_extractions_are_not_coverage_proof(self):
        for path in [self.before_files, self.after_files]:
            path.write_text('file,lines\n', encoding="utf-8")
        self.assertFalse(self.result()["equal"])

    def test_empty_or_failed_analysis_is_rejected(self):
        for data in [{"runs": []}, {"runs": [dict(self.sarif["runs"][0], invocations=[{"executionSuccessful": False}])]}]:
            with self.subTest(data=data):
                self.after_sarif.write_text(json.dumps(data), encoding="utf-8")
                with self.assertRaises(ValueError):
                    self.result()

    def test_equally_incomplete_query_sets_are_not_full_suite_proof(self):
        self.expected_rules.write_text('["cs/path-injection", "cs/sql-injection"]', encoding="utf-8")
        self.assertFalse(self.result()["equal"])


if __name__ == "__main__":
    unittest.main()
