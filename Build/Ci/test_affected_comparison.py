"""Protect report interpretation and failure handling, independently of dependency detection."""
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch


spec = importlib.util.spec_from_file_location("affected_comparison", Path(__file__).with_name("compare_affected.py"))
comparison = importlib.util.module_from_spec(spec)
spec.loader.exec_module(comparison)


class ComparisonTests(unittest.TestCase):
    def test_discovery_omissions_are_not_counted_as_savings(self):
        result = comparison.differences({"Core.csproj", "Native.csproj"},
                                        {"Core.csproj", "Example.csproj"},
                                        {"Core.csproj", "Example.csproj"})
        self.assertEqual(["Native.csproj"], result["currentOutsideDiscovery"])
        self.assertEqual([], result["currentOnlyReviewCandidates"])
        self.assertEqual(["Example.csproj"], result["affectedOnlyDependencyCandidates"])
        self.assertEqual(-1, result["potentialSavings"]["netProjectReduction"])
        self.assertEqual(0, result["potentialSavings"]["candidateRemovalPercent"])

    def test_inconsistent_observations_cannot_claim_savings(self):
        result = comparison.differences({"Core.csproj", "Tests.csproj"},
                                        {"Core.csproj"}, {"Core.csproj", "Tests.csproj"}, stable=False)
        self.assertEqual(["Tests.csproj"], result["currentOnlyReviewCandidates"])
        self.assertIsNone(result["potentialSavings"])
        with self.assertRaisesRegex(ValueError, "discovery inventory"):
            comparison.differences(set(), {"Unknown.csproj"}, set())

    def test_framework_duplicates_collapse_but_foreign_or_invalid_paths_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            project = root / "Core.csproj"
            project.write_text("<Project />")
            output = root / "result.json"
            row = {"Name": "Core", "FilePath": str(project)}
            output.write_text(json.dumps([row, row]))
            self.assertEqual({"Core.csproj"}, comparison.read_projects(output, root))
            for invalid in [{}, {"FilePath": "Core.csproj"}, {"FilePath": str(root.parent / "Foreign.csproj")}]:
                with self.subTest(invalid=invalid):
                    output.write_text(json.dumps([invalid]))
                    with self.assertRaises(ValueError):
                        comparison.read_projects(output, root)

    def test_tool_errors_or_missing_output_cannot_reuse_a_stale_result(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            # A previous successful report is deliberately present outside the new invocation directory.
            (root / "affected.json").write_text("[]")
            for code in (1, 0):
                with self.subTest(exit_code=code):
                    with patch.object(comparison.subprocess, "run", return_value=subprocess.CompletedProcess([], code)):
                        with self.assertRaises((RuntimeError, FileNotFoundError)):
                            comparison.run_tool("tool", root, "base", None, root, "affected", 5)
            with patch.object(comparison.subprocess, "run", return_value=subprocess.CompletedProcess([], 166)):
                projects, _ = comparison.run_tool("tool", root, "base", None, root, "affected", 5)
                self.assertEqual(set(), projects)


if __name__ == "__main__":
    unittest.main()
