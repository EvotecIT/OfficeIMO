"""Reachable selection contracts: transitive consumers, shared inputs and full CI."""
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

spec = importlib.util.spec_from_file_location("ci_selection", Path(__file__).with_name("select.py"))
selection = importlib.util.module_from_spec(spec)
spec.loader.exec_module(selection)


class SelectionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name).resolve()
        subprocess.run(["git", "init", "--quiet", str(self.root)], check=True)
        self.project("OfficeIMO.Core")
        self.project("OfficeIMO.Word", "OfficeIMO.Core")
        self.project("OfficeIMO.Word.Tests", "OfficeIMO.Word")
        self.project("OfficeIMO.CSV")
        self.project("OfficeIMO.CSV.Tests", "OfficeIMO.CSV")
        config = {
            "suites": {"word": ["OfficeIMO.Word.Tests"], "csv": ["OfficeIMO.CSV.Tests"],
                       "browser-exports": [], "website": [], "powershell": []},
            "partitions": {"word-rtf-html": ["OfficeIMO.Word.Tests"], "other-projects": ["OfficeIMO.CSV.Tests"]},
            "windows": [{"name": "Word", "project": "OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj"},
                        {"name": "CSV", "project": "OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj"}],
        }
        path = self.root / "Build/Ci/pr-suites.json"
        path.parent.mkdir(parents=True)
        path.write_text(json.dumps(config))
        self.stage()

    def project(self, name, dependency=None):
        path = self.root / name / (Path(name).name + ".csproj")
        path.parent.mkdir(parents=True, exist_ok=True)
        reference = f'<ItemGroup><ProjectReference Include="../{dependency}/{dependency}.csproj" /></ItemGroup>' if dependency else ""
        path.write_text(f'<Project Sdk="Microsoft.NET.Sdk">{reference}</Project>')

    def stage(self):
        subprocess.run(["git", "add", "."], cwd=self.root, check=True)

    def choose(self, *paths, full=False):
        return selection.select(self.root, list(paths), full)

    def test_transitive_consumers_run_without_unrelated_csv(self):
        result = self.choose("OfficeIMO.Core/Document.cs")
        self.assertTrue(result["suites"]["word"])
        self.assertFalse(result["suites"]["csv"])
        self.assertIn("OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj", result["projects"])

    def test_test_only_change_runs_its_own_contract(self):
        result = self.choose("OfficeIMO.Word.Tests/SaveTests.cs")
        self.assertTrue(result["suites"]["word"])
        self.assertFalse(result["suites"]["csv"])

    def test_pr_uses_latest_linux_tfm_and_retains_windows_contract(self):
        result = self.choose("OfficeIMO.Word/Save.cs")
        self.assertEqual(["net10.0"], [r["framework"] for r in result["ubuntu"]["include"]])
        self.assertEqual(["Word"], [r["name"] for r in result["windows"]["include"]])

    def test_full_run_restores_every_partition_and_both_modern_tfms(self):
        result = self.choose(full=True)
        self.assertTrue(result["full"])
        self.assertEqual(4, len(result["ubuntu"]["include"]))
        self.assertEqual({"net8.0", "net10.0"}, {r["framework"] for r in result["ubuntu"]["include"]})
        self.assertEqual(2, len(result["windows"]["include"]))

    def test_shared_source_and_fixtures_fail_open(self):
        for path in ["OfficeIMO.SharedSource/Xml.cs", "OfficeIMO.TestAssets/image.png", "Directory.Build.props", ".github/workflows/dotnet-tests.yml"]:
            with self.subTest(path=path):
                result = self.choose(path)
                self.assertTrue(result["suites"]["word"])
                self.assertTrue(result["suites"]["csv"])
                self.assertFalse(result["full"])

    def test_deleted_or_new_unmapped_project_fails_open(self):
        result = self.choose("OfficeIMO.Unknown/OfficeIMO.Unknown.csproj")
        self.assertTrue(result["suites"]["word"])
        self.assertTrue(result["suites"]["csv"])

    def test_tracked_unmapped_test_project_requires_a_route(self):
        self.project("OfficeIMO.New.Tests")
        self.stage()
        with self.assertRaisesRegex(ValueError, "OfficeIMO.New.Tests"):
            self.choose("OfficeIMO.New.Tests/OfficeIMO.New.Tests.csproj")

    def test_library_without_a_named_suite_selects_the_build_lane(self):
        self.project("OfficeIMO.New")
        self.stage()
        result = self.choose("OfficeIMO.New/OfficeIMO.New.csproj")
        self.assertTrue(result["suites"]["build-matrix"])
        self.assertIn("OfficeIMO.New/OfficeIMO.New.csproj", result["projects"])

    def test_executable_markdown_fixtures_select_contracts(self):
        for path in ["OfficeIMO.TestAssets/Documents/MarkdownVisuals/chart-summary.md",
                     "OfficeIMO.Word.Tests/Fixtures/profile.md"]:
            with self.subTest(path=path):
                result = self.choose(path)
                self.assertTrue(result["suites"]["word"])
                self.assertTrue(result["suites"]["csv"])

    def test_linked_compile_and_resource_consumers_run(self):
        path = self.root / "OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj"
        for kind in ["Compile", "None", "EmbeddedResource", "Content", "AdditionalFiles"]:
            with self.subTest(kind=kind):
                path.write_text(f'<Project><ItemGroup><{kind} Include="../OfficeIMO.Word/Shared*.*" /></ItemGroup></Project>')
                result = self.choose("OfficeIMO.Word/SharedInput.cs")
                self.assertTrue(result["suites"]["word"])
                self.assertTrue(result["suites"]["csv"])

    def test_docs_only_has_valid_but_disabled_matrices(self):
        result = self.choose("README.md", "Docs/guide.md")
        self.assertFalse(result["managed"])
        self.assertFalse(result["ubuntu_enabled"])
        self.assertFalse(result["windows_enabled"])
        self.assertEqual("none", result["ubuntu"]["include"][0]["partition"])

    def test_website_only_selects_website_without_library_tests(self):
        result = self.choose("Website/static/js/app.js")
        self.assertTrue(result["suites"]["website"])
        self.assertFalse(result["managed"])

    def test_javascript_only_selects_browser_exports_without_library_tests(self):
        result = self.choose("OfficeIMO.JavaScript/src/export.ts")
        self.assertTrue(result["suites"]["browser-exports"])
        self.assertTrue(result["suites"]["website"])
        self.assertFalse(result["managed"])

    def test_website_metadata_keeps_roadmap_and_powershell_inputs(self):
        for path in ["Docs/ROADMAP.md", "PowerShell/PSWriteOffice/Docs/command.md"]:
            with self.subTest(path=path):
                self.assertTrue(self.choose(path)["suites"]["website"])

    def test_unowned_powershell_assets_keep_their_suite_without_all_managed_tests(self):
        for path in ["PowerShell/PSWriteOffice/PSWriteOffice.Tests.ps1",
                     "PowerShell/PSWriteOffice/PSWriteOffice.psm1",
                     "PowerShell/PSWriteOffice/Docs/command.md"]:
            with self.subTest(path=path):
                result = self.choose(path)
                self.assertTrue(result["suites"]["powershell"])
                self.assertTrue(result["suites"]["website"])
                self.assertFalse(result["managed"])

    def test_linked_powershell_asset_still_selects_its_managed_consumer(self):
        path = self.root / "OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj"
        path.write_text('<Project><ItemGroup><Content Include="../PowerShell/PSWriteOffice/*.ps1" /></ItemGroup></Project>')
        result = self.choose("PowerShell/PSWriteOffice/PSWriteOffice.Tests.ps1")
        self.assertTrue(result["suites"]["powershell"])
        self.assertTrue(result["suites"]["csv"])
        self.assertFalse(result["suites"]["word"])

    def test_conditional_references_are_not_lost_on_another_os(self):
        path = self.root / "OfficeIMO.Word/OfficeIMO.Word.csproj"
        path.write_text('<Project><ItemGroup Condition="false"><ProjectReference Include="../OfficeIMO.CSV/OfficeIMO.CSV.csproj" /></ItemGroup></Project>')
        result = self.choose("OfficeIMO.CSV/Read.cs")
        self.assertTrue(result["suites"]["word"])
        self.assertTrue(result["suites"]["csv"])

    def test_powershell_root_property_keeps_library_consumers(self):
        path = self.root / "PowerShell/PSWriteOffice/Sources/PSWriteOffice/PSWriteOffice.csproj"
        path.parent.mkdir(parents=True)
        path.write_text('<Project><ItemGroup><ProjectReference Include="$(OfficeIMORoot)\\OfficeIMO.Word\\OfficeIMO.Word.csproj" /></ItemGroup></Project>')
        self.stage()
        result = self.choose("OfficeIMO.Word/Save.cs")
        self.assertIn(path.relative_to(self.root).as_posix(), result["projects"])


if __name__ == "__main__":
    unittest.main()
