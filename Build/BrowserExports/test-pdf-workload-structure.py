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


def fixture(wrong_owner=False, wrong_header=False, wrong_scope=False, *, integer_k=False,
            split_cell=False, missing_default_page=False, wrong_default_page=False,
            wrong_later_page=False):
    headers, values = verifier.HEADERS[:6], verifier.expected_row(0, 6)

    def painted_mark(mcid, text, role=None):
        return "/" + (role or ("TH" if mcid < 6 else "TD")) + " << /MCID " + str(mcid) + ">> BDC\nBT /F1 8 Tf 1 0 0 1 20 20 Tm (" + text + ") Tj ET\nEMC\n"

    def logical_mark(mcid, page=3):
        if integer_k and page == 3:
            return str(mcid)
        return "<< /Type /MCR /Pg " + str(page) + " 0 R /MCID " + str(mcid) + " >>"

    def default_page(first_data_cell=False):
        if not integer_k or first_data_cell and missing_default_page:
            return ""
        page = (26 if split_cell else 5) if first_data_cell and wrong_default_page else 3
        return " /Pg " + str(page) + " 0 R"

    content = "".join(painted_mark(index, text[:2] if split_cell and index == 6 else text)
                      for index, text in enumerate(headers + values))
    if split_cell:
        content += painted_mark(12, values[0][4:])
    owners = [18 if wrong_owner and index == 6 else (10 + index if index < 6 else 17 + index - 6)
              for index in range(12)]
    if split_cell:
        owners.append(17)
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
        23: "[" + " ".join(str(owner) + " 0 R" for owner in owners) + "]",
        24: "<< /Nums [0 23 0 R] >>",
        25: "<< /Names [" + " ".join("(h" + str(column) + ") " + str(10 + column) + " 0 R" for column in range(6)) + "] >>"
    }
    for column in range(6):
        scope = "/Row" if wrong_scope and column == 0 else "/Column"
        bodies[10 + column] = "<< /Type /StructElem /S /TH /P 9 0 R /ID (h" + str(column) + ") /A << /O /Table /Scope " + scope + " >>" + default_page() + " /K [" + logical_mark(column) + "] >>"
        children = [logical_mark(6 + column)]
        if split_cell and column == 0:
            # Return to the default page after the explicit later-page fragment.
            later_page = 3 if wrong_later_page else 26
            children.extend(["<< /Type /MCR /Pg " + str(later_page) + " 0 R /MCID 0 >>", logical_mark(12)])
        bodies[17 + column] = "<< /Type /StructElem /S /TD /P 16 0 R /A << /O /Table /Headers [(h" + str(1 if wrong_header and column == 0 else column) + ")] >>" + default_page(column == 0) + " /K [" + " ".join(children) + "] >>"
    if split_cell:
        later_content = painted_mark(0, values[0][2:4], role="TD")
        bodies[2] = "<< /Type /Pages /Count 2 /Kids [3 0 R 26 0 R] >>"
        bodies[24] = "<< /Nums [0 23 0 R 1 28 0 R] >>"
        bodies[26] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 27 0 R /StructParents 1 /Tabs /S >>"
        bodies[27] = "<< /Length " + str(len(later_content.encode())) + " >>\nstream\n" + later_content + "\nendstream"
        bodies[28] = "[17 0 R]"
    size = len(bodies) + 1
    output, offsets = bytearray(b"%PDF-1.7\n"), [0]
    for identity in range(1, size):
        offsets.append(len(output)); output.extend((str(identity) + " 0 obj\n" + bodies[identity] + "\nendobj\n").encode())
    start = len(output)
    output.extend(("xref\n0 " + str(size) + "\n0000000000 65535 f \n" + "".join(str(offset).zfill(10) + " 00000 n \n" for offset in offsets[1:]) +
                   "trailer\n<< /Size " + str(size) + " /Root 1 0 R >>\nstartxref\n" + str(start) + "\n%%EOF\n").encode())
    return PdfReader(BytesIO(output), strict=True)


class StructureVerifierTests(unittest.TestCase):
    def test_real_fixture_counts_values_headers_and_unique_owners(self):
        digests = []
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k):
                proof = verifier.verify_reader(fixture(integer_k=integer_k), 1, 6)
                self.assertEqual((proof["dataCells"], proof["markedFragments"], proof["logicalRows"]), (6, 12, 1))
                self.assertTrue(proof["parentTree"] and proof["headerAssociations"] and proof["uniqueMarkedOwners"])
                digests.append(proof["logicalSha256"])
        self.assertEqual(digests[0], digests[1])

    def test_split_cell_keeps_value_and_owners_across_pages(self):
        expected = verifier.verify_reader(fixture(), 1, 6)["logicalSha256"]
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k):
                proof = verifier.verify_reader(fixture(integer_k=integer_k, split_cell=True), 1, 6)
                self.assertEqual((proof["dataCells"], proof["markedFragments"], proof["logicalRows"]), (6, 14, 1))
                self.assertEqual(proof["logicalSha256"], expected)
                self.assertTrue(proof["parentTree"] and proof["uniqueMarkedOwners"])

    def test_wrong_count_is_rejected(self):
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k), self.assertRaisesRegex(ValueError, "counts differ"):
                verifier.verify_reader(fixture(integer_k=integer_k), 2, 6)

    def test_wrong_parent_tree_owner_is_rejected(self):
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k), self.assertRaisesRegex(ValueError, "ParentTree owner differs"):
                verifier.verify_reader(fixture(wrong_owner=True, integer_k=integer_k), 1, 6)

    def test_wrong_column_header_identity_is_rejected(self):
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k), self.assertRaisesRegex(ValueError, "header association differs"):
                verifier.verify_reader(fixture(wrong_header=True, integer_k=integer_k), 1, 6)

    def test_row_scope_on_a_column_heading_is_rejected(self):
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k), self.assertRaisesRegex(ValueError, "Column heading scope differs"):
                verifier.verify_reader(fixture(wrong_scope=True, integer_k=integer_k), 1, 6)

    def test_integer_mark_without_default_page_is_rejected(self):
        for split_cell in (False, True):
            with self.subTest(split_cell=split_cell), self.assertRaisesRegex(ValueError, "Logical mark has no physical content"):
                verifier.verify_reader(fixture(integer_k=True, split_cell=split_cell, missing_default_page=True), 1, 6)

    def test_integer_mark_with_wrong_default_page_is_rejected(self):
        for split_cell in (False, True):
            with self.subTest(split_cell=split_cell), self.assertRaisesRegex(ValueError, "Logical mark has no physical content"):
                verifier.verify_reader(fixture(integer_k=True, split_cell=split_cell, wrong_default_page=True), 1, 6)

    def test_later_mcr_with_wrong_page_is_rejected(self):
        for integer_k in (False, True):
            with self.subTest(integer_k=integer_k), self.assertRaisesRegex(ValueError, "duplicate logical owners"):
                verifier.verify_reader(fixture(integer_k=integer_k, split_cell=True, wrong_later_page=True), 1, 6)


if __name__ == "__main__":
    unittest.main()
