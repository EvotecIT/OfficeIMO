"""Scheduling, native failure isolation and the actual evidence-comparison boundary."""
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

import compare_netframework_shards as evidence

spec = importlib.util.spec_from_file_location("ci_selection", Path(__file__).with_name("select.py"))
selection = importlib.util.module_from_spec(spec)
spec.loader.exec_module(selection)


def contract(name):
    return {"name": name, "project": f"OfficeIMO.{name}.Tests/OfficeIMO.{name}.Tests.csproj"}


class ShardTests(unittest.TestCase):
    def test_full_inventory_is_covered_once_in_eleven_balanced_shards(self):
        config = json.loads(Path(__file__).with_name("pr-suites.json").read_text())
        rows = config["windows"]
        shards = selection.windows_shards(rows, config["windows_shard_seconds"])
        self.assertEqual(11, len(shards))
        self.assertCountEqual(rows, [r for s in shards for r in s["projects"]])
        loads = [s["estimated_seconds"] for s in shards]
        self.assertLessEqual(max(loads) - min(loads), 60)
        self.assertEqual(shards, selection.windows_shards(list(reversed(rows)), config["windows_shard_seconds"]))

    def test_small_selection_and_unmeasured_new_project_keep_every_contract(self):
        rows = [contract("New"), contract("Word")]
        shards = selection.windows_shards(rows, {"Word": 434})
        self.assertEqual(2, len(shards))
        self.assertCountEqual(rows, [r for s in shards for r in s["projects"]])
        self.assertEqual([434, 180], [s["estimated_seconds"] for s in shards])

    def test_duplicate_routes_and_invalid_estimates_fail_selection(self):
        for rows, weights in [([contract("Word")] * 2, {}),
                              ([contract("Word")], {"Word": -1}),
                              ([contract("Word")], {"Word": float("nan")})]:
            with self.subTest(rows=rows, weights=weights), self.assertRaises(ValueError):
                selection.windows_shards(rows, weights)

    def test_native_failures_do_not_stop_later_projects_or_remaining_packs(self):
        runner = Path(__file__).with_name("Invoke-NetFrameworkShard.ps1").resolve()
        pwsh = shutil.which("pwsh")
        self.assertIsNotNone(pwsh, "The CI shard contract requires PowerShell 7")
        for failure in ("restore", "test", "pack"):
            with self.subTest(failure=failure), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                fake = root / "fake.py"
                fake.write_text('''import json, os, sys
from pathlib import Path
args = sys.argv[1:]
with open(os.environ["CALL_LOG"], "a") as output:
    output.write(json.dumps(args) + "\\n")
fail = os.environ["FAIL_STEP"]
sys.exit(1 if (args[0] == fail and ("OfficeIMO.Fail." in args[1] or "OfficeIMO.Excel.CodeGlyphX/" in args[1])) else 0)
''')
                if os.name == "nt":
                    # A native exe exercises the same argument passing as dotnet.exe;
                    # cmd wrappers reinterpret the filter's ampersand as shell syntax.
                    source = root / "FakeDotnet.cs"
                    source.write_text(r'''using System;
using System.IO;
class FakeDotnet {
    static int Main(string[] args) {
        string json = "[" + String.Join(",", Array.ConvertAll(args, s => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"")) + "]\n";
        File.AppendAllText(Environment.GetEnvironmentVariable("CALL_LOG"), json);
        return args[0] == Environment.GetEnvironmentVariable("FAIL_STEP") &&
            (args[1].Contains("OfficeIMO.Fail.") || args[1].Contains("OfficeIMO.Excel.CodeGlyphX/")) ? 1 : 0;
    }
}''')
                    compiler = Path(os.environ["WINDIR"]) / "Microsoft.NET/Framework64/v4.0.30319/csc.exe"
                    subprocess.run([str(compiler), "/nologo", f"/out:{root / 'dotnet.exe'}", str(source)], check=True,
                                   capture_output=True, text=True)
                else:
                    command = root / "dotnet"
                    command.write_text(f'#!{sys.executable}\nexec(compile(open({str(fake)!r}).read(), {str(fake)!r}, "exec"))\n')
                    command.chmod(0o755)
                env = dict(os.environ, PATH=str(root) + os.pathsep + os.environ["PATH"],
                           CALL_LOG=str(root / "calls.jsonl"), FAIL_STEP=failure)
                rows = [contract("Fail"), contract("After"),
                        {"name": "Bridge", "project": "OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj"}]
                result = subprocess.run([pwsh, "-NoProfile", "-File", str(runner), "-ProjectsJson", json.dumps(rows),
                                         "-EvidenceDirectory", str(root / "evidence"), "-DiagnosticsDirectory", str(root / "logs")],
                                        cwd=root, env=env, capture_output=True, text=True)
                self.assertEqual(1, result.returncode, result.stdout + result.stderr)
                calls = [json.loads(line) for line in (root / "calls.jsonl").read_text().splitlines()]
                tests = [c for c in calls if c[0] == "test"]
                self.assertIn(rows[1]["project"], [c[1] for c in tests])
                for call in tests:
                    self.assertEqual("net472", call[call.index("--framework") + 1])
                    self.assertEqual("Category!=ExcelImageVisualGate&Category!=Performance", call[call.index("--filter") + 1])
                    self.assertIn("--no-restore", call)
                self.assertEqual(5, len([c for c in calls if c[0] == "pack"]))
                records = json.loads((root / "evidence/execution.json").read_text(encoding="utf-8-sig"))
                self.assertEqual([r["project"] for r in rows], [r["project"] for r in records])
                self.assertEqual(1, sum(s["exit_code"] != 0 for r in records for s in r["steps"]))


class EvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.row = contract("Word")
        for layout in ("baseline", "sharded"):
            self.fixture(layout)

    def fixture(self, layout, test="Save", outcome="Passed", row=None):
        row = row or self.row
        directory = self.root / f"net472-evidence-1-{layout}-Word"
        directory.mkdir(exist_ok=True)
        (directory / "environment.json").write_text(json.dumps({"source": "same-revision", "sdk": "10.0.100",
                                                                "image": "same-image", "arch": "X64"}))
        steps = [{"name": f'{row["name"]} / {name}', "exit_code": 0, "seconds": 10} for name in ("restore", "net472")]
        if row["project"] == "OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj":
            steps += [{"name": f'{row["name"]} / pack {p}', "exit_code": 0, "seconds": 10}
                      for p in ("Drawing", "Word", "Excel", "PowerPoint", "Pdf")]
        (directory / "execution.json").write_text(json.dumps([dict(row, steps=steps)]))
        assembly = Path(row["project"]).stem
        (directory / "tests.trx").write_text(f'''<TestRun xmlns="{evidence.NS['t']}">
<TestDefinitions><UnitTest id="1"><TestMethod codeBase="C:\\bin\\{assembly}.dll" className="Word.Tests" name="{test}"/></UnitTest></TestDefinitions>
<Results><UnitTestResult testId="1" testName="{test}(input: 1)" outcome="{outcome}"/></Results>
<ResultSummary><Counters total="1"/></ResultSummary></TestRun>''')
        return directory

    def test_equal_test_identities_and_outcomes_pass_with_complete_evidence(self):
        report = evidence.compare(self.root, [self.row], rounds=1)
        self.assertEqual([1, 1], [r["test_results"] for r in report])
        self.assertEqual([20, 20], [r["command_seconds"] for r in report])

    def test_forged_contract_name_cannot_hide_required_pack_records(self):
        row = {"name": "Drawing-CodeGlyphX", "project": "OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj"}
        self.fixture("baseline", row=row)
        directory = self.fixture("sharded", row=row)
        path = directory / "execution.json"
        records = json.loads(path.read_text())
        records[0]["name"] = "Other"
        records[0]["steps"] = [{"name": f"Other / {s}", "seconds": 10, "exit_code": 0} for s in ("restore", "net472")]
        path.write_text(json.dumps(records))
        with self.assertRaisesRegex(ValueError, "mismatched contract name"):
            evidence.compare(self.root, [row], rounds=1)

    def test_same_basename_in_an_unselected_directory_cannot_satisfy_inventory(self):
        directory = self.root / "net472-evidence-1-sharded-Word"
        path = directory / "execution.json"
        records = json.loads(path.read_text())
        records[0]["project"] = "OfficeIMO.Unselected/OfficeIMO.Word.Tests.csproj"
        path.write_text(json.dumps(records))
        with self.assertRaisesRegex(ValueError, "Unknown project"):
            evidence.compare(self.root, [self.row], rounds=1)

    def test_configured_display_name_change_retains_project_pack_requirements(self):
        row = {"name": "Bridge", "project": "OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj"}
        self.fixture("baseline", row=row)
        self.fixture("sharded", row=row)
        report = evidence.compare(self.root, [row], rounds=1)
        self.assertEqual([70, 70], [r["command_seconds"] for r in report])

    def test_missing_codeglyphx_pack_record_fails_the_canonical_contract(self):
        row = {"name": "Drawing-CodeGlyphX", "project": "OfficeIMO.Drawing.CodeGlyphX.Tests/OfficeIMO.Drawing.CodeGlyphX.Tests.csproj"}
        self.fixture("baseline", row=row)
        directory = self.fixture("sharded", row=row)
        path = directory / "execution.json"
        records = json.loads(path.read_text())
        records[0]["steps"].pop()
        path.write_text(json.dumps(records))
        with self.assertRaisesRegex(ValueError, "execution obligations"):
            evidence.compare(self.root, [row], rounds=1)

    def test_equal_counts_with_changed_test_identity_or_skip_fail_parity(self):
        for test, outcome in (("Other", "Passed"), ("Save", "NotExecuted")):
            with self.subTest(test=test, outcome=outcome):
                self.fixture("sharded", test, outcome)
                with self.assertRaisesRegex(ValueError, "parity failed"):
                    evidence.compare(self.root, [self.row], rounds=1)

    def test_changed_sdk_or_source_cannot_be_reported_as_matched_measurements(self):
        for field in ("source", "sdk", "image", "arch"):
            with self.subTest(field=field):
                directory = self.fixture("sharded")
                path = directory / "environment.json"
                metadata = json.loads(path.read_text())
                metadata[field] = "different"
                path.write_text(json.dumps(metadata))
                with self.assertRaisesRegex(ValueError, "environment differs"):
                    evidence.compare(self.root, [self.row], rounds=1)

    def test_missing_duplicate_or_failed_execution_evidence_is_rejected(self):
        directory = self.root / "net472-evidence-1-sharded-Word"
        for problem in ("missing-trx", "duplicate-project", "failed-command", "missing-test-command"):
            with self.subTest(problem=problem):
                self.fixture("sharded")
                manifest = directory / "execution.json"
                records = json.loads(manifest.read_text())
                if problem == "missing-trx":
                    (directory / "tests.trx").unlink()
                elif problem == "duplicate-project":
                    records += records
                elif problem == "failed-command":
                    records[0]["steps"][1]["exit_code"] = 1
                else:
                    records[0]["steps"].pop()
                manifest.write_text(json.dumps(records))
                with self.assertRaises(ValueError):
                    evidence.compare(self.root, [self.row], rounds=1)


if __name__ == "__main__":
    unittest.main()
