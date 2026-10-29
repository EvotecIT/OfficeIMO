"""Exercise the independent artifact verifier with valid and broken package relationships."""

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from verify import verify


class ArtifactVerificationTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.word = {"word/document.xml": '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Quarterly report</w:t></w:r></w:p><w:p><w:r><w:t>Created with OfficeIMO.Word.</w:t></w:r></w:p>
<w:tbl><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Area</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Documents</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Generated</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''}
        self.excel = {
            "xl/workbook.xml": '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" r:id="sheetData"/></sheets></workbook>''',
            "xl/_rels/workbook.xml.rels": self.relationship("sheetData", "worksheet", "worksheets/data.xml"),
            "xl/worksheets/data.xml": '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row>
<c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>42</v></c>
</row></sheetData><tableParts><tablePart r:id="tableData"/></tableParts></worksheet>''',
            "xl/worksheets/_rels/data.xml.rels": self.relationship("tableData", "table", "../tables/data.xml"),
            "xl/tables/data.xml": '<table name="DataTable" ref="A1:B2"/>',
            "xl/tables/orphan.xml": '<table name="DataTable" ref="A1:B2"/>',
        }

    @staticmethod
    def relationship(identifier, kind, target):
        return ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="{identifier}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>'
                '</Relationships>')

    def write_packages(self):
        for filename, parts in (("report.docx", self.word), ("report.xlsx", self.excel)):
            with ZipFile(self.root / filename, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)

    def test_valid_nondefault_worksheet_and_table_paths(self):
        self.write_packages()
        verify(self.root)

    def test_disabled_first_row_header_is_rejected(self):
        original = self.word["word/document.xml"]
        for disabled in ("0", "false", "off"):
            with self.subTest(disabled=disabled):
                self.word["word/document.xml"] = original.replace('<w:tblHeader/>', f'<w:tblHeader w:val="{disabled}"/>')
                self.write_packages()
                with self.assertRaisesRegex(ValueError, "header row"):
                    verify(self.root)

    def test_header_only_on_second_row_is_rejected(self):
        original = self.word["word/document.xml"].replace('<w:trPr><w:tblHeader/></w:trPr>', '')
        self.word["word/document.xml"] = original.replace('<w:tr><w:tc><w:p><w:r><w:t>Documents', '<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Documents')
        self.write_packages()
        with self.assertRaisesRegex(ValueError, "header row"):
            verify(self.root)

    def test_orphan_table_does_not_satisfy_missing_table_reference(self):
        self.excel["xl/worksheets/data.xml"] = self.excel["xl/worksheets/data.xml"].replace('<tableParts><tablePart r:id="tableData"/></tableParts>', '')
        self.write_packages()
        with self.assertRaisesRegex(ValueError, "DataTable"):
            verify(self.root)

    def test_orphan_correct_table_does_not_hide_wrong_referenced_table(self):
        self.excel["xl/tables/data.xml"] = '<table name="WrongTable" ref="A1:C3"/>'
        self.write_packages()
        with self.assertRaisesRegex(ValueError, "DataTable"):
            verify(self.root)

    def test_missing_worksheet_relationship_is_rejected(self):
        self.excel["xl/_rels/workbook.xml.rels"] = self.relationship("other", "worksheet", "worksheets/data.xml")
        self.write_packages()
        with self.assertRaisesRegex(ValueError, "worksheet relationship"):
            verify(self.root)


if __name__ == "__main__":
    unittest.main()
