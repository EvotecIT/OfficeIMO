"""Protect report interpretation and failure handling, independently of dependency detection."""
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
from types import SimpleNamespace
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

    def test_stable_omission_of_declared_consumers_also_withholds_savings(self):
        result = comparison.assess_observations(
            {"Core.csproj", "Tests.csproj"}, [{"Core.csproj"}] * 3,
            {"Core.csproj", "Tests.csproj"}, {"Tests.csproj": {"Core.csproj"}})
        self.assertEqual("dependency-gaps", result["status"])
        self.assertEqual([{"consumer": "Tests.csproj", "dependency": "Core.csproj"}],
                         result["missingDeclaredReferences"])
        self.assertIsNone(result["differences"]["potentialSavings"])

    def test_later_tool_failure_preserves_completed_observations(self):
        selector = SimpleNamespace(select=lambda *args: {"projects": ["Core.csproj"]},
                                   tracked=lambda *args: [], project_graph=lambda *args: ({}, {}))
        outputs = ["baseline", "revision", "", "", "Core/Input.cs\0"]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            with patch.object(comparison, "git", side_effect=outputs), \
                    patch.object(comparison, "load_selection", return_value=selector), \
                    patch.object(comparison.platform, "platform", return_value="TestOS"), \
                    patch.object(comparison.subprocess, "check_output", side_effect=["10.0.404", "7.0.1"]), \
                    patch.object(comparison, "run_tool", side_effect=[({"Core.csproj"}, .2),
                                                                     ({"Core.csproj"}, .3), RuntimeError("second evaluation failed")]):
                result = comparison.compare(root, "base", "tool", None, root, 5)
        self.assertEqual("incomplete", result["status"])
        self.assertEqual([["Core.csproj"]], result["observations"])
        self.assertEqual([.3], result["seconds"]["dotnetAffected"])
        self.assertEqual("second evaluation failed", result["error"])
        self.assertNotIn("differences", result)

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
