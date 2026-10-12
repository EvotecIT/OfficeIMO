"""Decode the controlled normal/array future-function profile with Python's ZIP/XML readers."""
import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
FUNCTION = re.compile(r"(?i)(?<![\w.])_xlfn\.(IFNA|MINIFS|MAXIFS)(?=\()")
EXPECTED = ["99", "#DIV/0!", "10", "30", "0", "0", "10", "30"]
EDITED = ["99", "#DIV/0!", "30", "50", "0", "0", "30", "50"]


def decode(path):
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
    cells = {}
    for cell in root.findall(".//s:c", NS):
        formula = cell.find("s:f", NS)
        cells[cell.attrib["r"]] = {
            "kind": cell.get("t", "n"), "value": cell.findtext("s:v", None, NS),
            "formula": formula.text if formula is not None else None,
            "formulaType": formula.get("t") if formula is not None else None,
            "formulaRange": formula.get("ref") if formula is not None else None,
        }
        if formula is not None:
            assert formula.text and formula.get("t") != "shared", "Profile requires explicit formulas"
    return cells


def canonical_case(formula):
    # The eight controlled expressions contain only the EU/missing string criteria.
    # Normalize these three function identifiers, never arbitrary formula text.
    return FUNCTION.sub(lambda match: "_xlfn." + match[1].upper(), formula)


def artifact(path):
    return {"path": path.name, "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ["bare_input", "bare_native", "qualified_input", "native", "edited_native", "output"]:
        parser.add_argument(name, type=Path)
    parser.add_argument("--source-revision", required=True)
    parser.add_argument("--storage-source-sha256", required=True)
    parser.add_argument("--artifacts", nargs="*", type=Path, default=[])
    args = parser.parse_args()
    assert re.fullmatch(r"[0-9a-f]{40}", args.source_revision)
    assert re.fullmatch(r"[0-9a-f]{64}", args.storage_source_sha256)
    bare, negative, source, native, edited = [decode(path) for path in
        [args.bare_input, args.bare_native, args.qualified_input, args.native, args.edited_native]]
    assert source["B2"]["value"] == native["B2"]["value"] == "10"
    assert edited["B2"]["value"] == "50"
    cases = []
    for row in range(2, 10):
        for column, index in [("E", 5), ("F", 6), ("G", 7), ("H", 8)]:
            address = column + str(row)
            authored, reference, changed = source[address], native[address], edited[address]
            assert authored["formula"] and authored["value"] is None
            assert canonical_case(authored["formula"]) == reference["formula"] == changed["formula"]
            assert reference["value"] == EXPECTED[row - 2] and changed["value"] == EDITED[row - 2]
            expected_kind = "e" if row == 3 else "n"
            assert reference["kind"] == changed["kind"] == expected_kind
            expected_array = column in "GH"
            for cell in [authored, reference, changed]:
                assert (cell["formulaType"] == "array") == expected_array
            if column in "EG":
                assert "_xlfn." not in bare[address]["formula"]
                assert negative[address]["kind"] == "e" and negative[address]["value"] == "#NAME?"
            else:
                assert negative[address]["value"] == reference["value"]
                assert negative[address]["kind"] == reference["kind"]
            cases.append({"address": address, "row": row, "column": index,
                "formula": authored["formula"], "nativeFormula": reference["formula"],
                "formulaType": authored["formulaType"], "formulaRange": authored["formulaRange"],
                "kind": reference["kind"], "value": reference["value"], "editedValue": changed["value"]})
    for cells in [bare, source, negative, native, edited]:
        assert sum(cell["formula"] is not None for cell in cells.values()) == 32
    inputs = [args.bare_input, args.bare_native, args.qualified_input, args.native,
              args.edited_native, *args.artifacts]
    result = {"producer": "Microsoft Excel 16.113.3 (16.113.26092714)",
        "decoder": "Python standard-library zipfile and ElementTree; no OfficeIMO decoder",
        "authoring": {"sourceRevision": args.source_revision,
            "owner": "OfficeIMO.Excel/ExcelFormulaStorage.cs", "ownerSha256": args.storage_source_sha256,
            "publicApis": ["CellFormula", "SetArrayFormula"], "cachedInputs": False},
        "specification": "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf",
        "functions": ["IFNA", "MINIFS", "MAXIFS"], "formulaCells": 32,
        "input": args.qualified_input.name, "native": args.native.name, "editedNative": args.edited_native.name,
        "edit": {"cell": "B2", "before": 10, "after": 50},
        "negativeControl": {"input": args.bare_input.name, "native": args.bare_native.name,
            "bareNormalNameErrors": 8, "bareArrayNameErrors": 8, "qualifiedControls": 16},
        "cells": cases, "inputs": [artifact(path) for path in inputs],
        "limits": ["Eight controlled expressions in normal and single-cell array formulas; typed errors and numeric criteria only.",
            "Native Excel canonicalizes lower-case function identifiers; authored case is recorded separately.",
            "Wider future functions, dynamic spills, other applications and producer versions retain separate qualification gates."]}
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"formulaCells": len(cases), "inputs": len(inputs), "bareNameErrors": 16}))


if __name__ == "__main__":
    main()
