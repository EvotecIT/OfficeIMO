"""Regression checks for Registry validation and packaged ownership metadata."""

import io
import unittest
from zipfile import ZipFile

from jsonschema import ValidationError

from validate_mcp_registry import (
    SCHEMA_URL, validate_marker, validate_metadata, validate_package,
)


class RegistryValidationTests(unittest.TestCase):
    def setUp(self):
        self.server = {"$schema": SCHEMA_URL, "name": "io.github.evotecit/officeimo"}
        self.marker = "<!-- mcp-name: io.github.evotecit/officeimo -->"

    def test_metadata_cannot_select_a_permissive_schema(self):
        server = dict(self.server, **{"$schema": "https://example.com/permissive.json"})
        with self.assertRaisesRegex(ValueError, "pinned official schema"):
            validate_metadata(server, {})

    def test_metadata_must_satisfy_the_selected_schema(self):
        with self.assertRaises(ValidationError):
            validate_metadata(self.server, {"required": ["version"]})

    def test_missing_mismatched_and_duplicate_markers_are_rejected(self):
        for readme in ("", "<!-- mcp-name: io.github.other/server -->",
                       self.marker + "\n" + self.marker):
            with self.subTest(readme=readme), self.assertRaises(ValueError):
                validate_marker(readme, self.server)

    def package(self, readme_element, readme):
        package = io.BytesIO()
        with ZipFile(package, "w") as archive:
            archive.writestr("OfficeIMO.Tool.nuspec", f'''
                <package xmlns="http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd">
                  <metadata><id>OfficeIMO.Tool</id>{readme_element}</metadata>
                </package>''')
            archive.writestr("docs/package-readme.md", readme)
        package.seek(0)
        return package

    def test_package_checks_its_declared_readme(self):
        validate_package(self.package("<readme>docs/package-readme.md</readme>",
                                      self.marker), self.server)
        with self.assertRaises(ValueError):
            validate_package(self.package("<readme>docs/package-readme.md</readme>",
                                          "README without ownership"), self.server)

    def test_package_requires_a_readme_declaration(self):
        with self.assertRaisesRegex(ValueError, "declare a README"):
            validate_package(self.package("", self.marker), self.server)


if __name__ == "__main__":
    unittest.main()
