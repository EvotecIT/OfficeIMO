import json
import tempfile
import unittest
import zipfile
from pathlib import Path

import validate_third_party_notices as checker

PROJECT = """<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup><PackageLicenseExpression>{expression}</PackageLicenseExpression></PropertyGroup>
  <ItemGroup>
    <None Include="THIRD-PARTY-NOTICES.md" Pack="true" PackagePath="\\" />
    <None Include="Licenses\\*.txt" Pack="true" PackagePath="Licenses/" />
  </ItemGroup>
</Project>
"""


class ThirdPartyNoticesTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.write("Lib/Lib.csproj", PROJECT.format(expression="MIT AND BSD-3-Clause"))
        self.write("Lib/Licenses/codec-LICENSE.txt", "BSD-3-Clause text")
        self.write("Lib/Licenses/data-LICENSE.txt", "Unicode License V3 text")
        self.write("Lib/THIRD-PARTY-NOTICES.md", "## Codec transform\n\n## Unicode tables\n")
        self.write("Lib/Transform.cs", "// Adapted from Codec (BSD-3-Clause).\nclass T {}\n")
        self.write("Lib/Tables.Generated.cs", "// https://www.unicode.org/Public/17.0.0/ucd/Scripts.txt\nclass U {}\n")
        self.write("Lib/Own.cs", "// Copyright (c) Evotec.\nclass O {}\n")
        self.manifest = {
            "entries": [
                self.entry("codec", "Codec transform", "code", "BSD-3-Clause", "Lib/Transform.cs", "codec-LICENSE.txt"),
                self.entry("unicode", "Unicode tables", "data", "Unicode-3.0", "Lib/Tables.Generated.cs", "data-LICENSE.txt"),
            ],
            "notices": {"Lib": {"path": "Lib/THIRD-PARTY-NOTICES.md", "packagePath": "THIRD-PARTY-NOTICES.md"}},
        }
        self.save_manifest()

    def tearDown(self):
        self.directory.cleanup()

    def entry(self, id, title, kind, license, file, license_file):
        return {
            "id": id, "title": title, "kind": kind, "license": license, "source": "https://example.invalid",
            "project": "Lib", "files": [file], "licenseFiles": {f"Lib/Licenses/{license_file}": f"Licenses/{license_file}"},
        }

    def write(self, name, text):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")

    def save_manifest(self):
        self.write(checker.MANIFEST, json.dumps(self.manifest))

    def errors(self):
        return checker.check_source(self.root).errors

    def test_declared_material_passes(self):
        self.assertEqual([], self.errors())

    def test_data_license_in_expression_is_rejected(self):
        self.write("Lib/Lib.csproj", PROJECT.format(expression="MIT AND BSD-3-Clause AND Unicode-3.0"))
        self.assertIn("without third-party code behind it: Unicode-3.0", "\n".join(self.errors()))

    def test_missing_code_license_in_expression_is_rejected(self):
        self.write("Lib/Lib.csproj", PROJECT.format(expression="MIT"))
        self.assertIn("missing: BSD-3-Clause", "\n".join(self.errors()))

    def test_undeclared_adapted_source_is_rejected(self):
        self.write("Lib/Other.cs", "// Ported from Elsewhere.\nclass X {}\n")
        self.assertIn("Lib/Other.cs:1: third-party marker", "\n".join(self.errors()))

    def test_package_without_license_text_is_rejected(self):
        package = self.root / "out" / "Lib.1.0.0.nupkg"
        package.parent.mkdir()
        with zipfile.ZipFile(package, "w") as archive:
            archive.writestr("Lib.nuspec", '<package xmlns="http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd"><metadata>'
                             '<id>Lib</id><license type="expression">MIT AND BSD-3-Clause</license></metadata></package>')
            archive.writestr("THIRD-PARTY-NOTICES.md", "## Codec transform\n\n## Unicode tables\n")
            archive.writestr("Licenses/codec-LICENSE.txt", "BSD-3-Clause text")
        errors = checker.check_packages(self.root, [package.parent]).errors
        self.assertEqual(["Lib.1.0.0.nupkg: missing 'Licenses/data-LICENSE.txt' for unicode."], errors)


if __name__ == "__main__":
    unittest.main()
