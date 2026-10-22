"""Opt-in verifier qualification; uses tiny in-memory real PDFs and the installed pypdf reader.

Run separately from ordinary product CI: python test-pdf-workload-structure.py
"""
import importlib.util
from io import BytesIO
from pathlib import Path
import unittest

from pypdf import PdfReader

spec = importlib.util.spec_from_file_location("workload_structure", Path(__file__).with_name("verify-pdf-workload-structure.py"))
verifier = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verifier)


def fixture(wrong_owner=False, wrong_header=False, wrong_scope=False):
    headers, values = verifier.HEADERS[:6], verifier.expected_row(0, 6)
    content = "".join("/" + ("TH" if index < 6 else "TD") + " << /MCID " + str(index) + ">> BDC\nBT /F1 8 Tf 1 0 0 1 20 20 Tm (" + text + ") Tj ET\nEMC\n"
                      for index, text in enumerate(headers + values))
    bodies = {
        1: "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /Lang (en-US) /StructTreeRoot 6 0 R >>",
        2: "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        3: "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /StructParents 0 /Tabs /S >>",
        4: "<< /Length " + str(len(content.encode())) + " >>\nstream\n" + content + "\nendstream",
        5: "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        6: "<< /Type /StructTreeRoot /K [7 0 R] /ParentTree 24 0 R /IDTree 25 0 R >>",
        7: "<< /Type /StructElem /S /Document /P 6 0 R /K [8 0 R] >>",
        8: "<< /Type /StructElem /S /Table /P 7 0 R /K [9 0 R 16 0 R] >>",
        9: "<< /Type /StructElem /S /TR /P 8 0 R /K [" + " ".join(str(10 + column) + " 0 R" for column in range(6)) + "] >>",
        16: "<< /Type /StructElem /S /TR /P 8 0 R /K [" + " ".join(str(17 + column) + " 0 R" for column in range(6)) + "] >>",
        23: "[" + " ".join(str(18 if wrong_owner and index == 6 else (10 + index if index < 6 else 17 + index - 6)) + " 0 R" for index in range(12)) + "]",
        24: "<< /Nums [0 23 0 R] >>",
        25: "<< /Names [" + " ".join("(h" + str(column) + ") " + str(10 + column) + " 0 R" for column in range(6)) + "] >>"
    }
    for column in range(6):
        scope = "/Row" if wrong_scope and column == 0 else "/Column"
        bodies[10 + column] = "<< /Type /StructElem /S /TH /P 9 0 R /ID (h" + str(column) + ") /A << /O /Table /Scope " + scope + " >> /K [<< /Type /MCR /Pg 3 0 R /MCID " + str(column) + " >>] >>"
        bodies[17 + column] = "<< /Type /StructElem /S /TD /P 16 0 R /A << /O /Table /Headers [(h" + str(1 if wrong_header and column == 0 else column) + ")] >> /K [<< /Type /MCR /Pg 3 0 R /MCID " + str(6 + column) + " >>] >>"
    output, offsets = bytearray(b"%PDF-1.7\n"), [0]
    for identity in range(1, 26):
        offsets.append(len(output)); output.extend((str(identity) + " 0 obj\n" + bodies[identity] + "\nendobj\n").encode())
    start = len(output)
    output.extend(("xref\n0 26\n0000000000 65535 f \n" + "".join(str(offset).zfill(10) + " 00000 n \n" for offset in offsets[1:]) +
                   "trailer\n<< /Size 26 /Root 1 0 R >>\nstartxref\n" + str(start) + "\n%%EOF\n").encode())
    return PdfReader(BytesIO(output), strict=True)


class StructureVerifierTests(unittest.TestCase):
    def test_real_fixture_counts_values_headers_and_unique_owners(self):
        proof = verifier.verify_reader(fixture(), 1, 6)
        self.assertEqual((proof["dataCells"], proof["markedFragments"], proof["logicalRows"]), (6, 12, 1))

    def test_wrong_count_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "counts differ"):
            verifier.verify_reader(fixture(), 2, 6)

    def test_wrong_parent_tree_owner_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "ParentTree owner differs"):
            verifier.verify_reader(fixture(wrong_owner=True), 1, 6)

    def test_wrong_column_header_identity_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "header association differs"):
            verifier.verify_reader(fixture(wrong_header=True), 1, 6)

    def test_row_scope_on_a_column_heading_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Column heading scope differs"):
            verifier.verify_reader(fixture(wrong_scope=True), 1, 6)


if __name__ == "__main__":
    unittest.main()
