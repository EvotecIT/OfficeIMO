"""Opt-in fixed second/millisecond duration evidence from unchanged Numbers and Apple XLSX files.

Run with the isolated numbers-parser 4.19.0 oracle. This script is not a product
runtime dependency and does not assess PDF appearance or native Excel rendering.
"""
import argparse
import hashlib
import json
import math
from importlib.metadata import version
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

from numbers_parser import Document
from numbers_parser.generated import TSKArchives_pb2

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("source", type=Path)
parser.add_argument("native_export", type=Path)
parser.add_argument("output", type=Path)
args = parser.parse_args()
if version("numbers-parser") != "4.19.0":
    raise RuntimeError("Use the pinned numbers-parser 4.19.0 evidence environment.")

document = Document(str(args.source))
ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
profiles = []
with zipfile.ZipFile(args.native_export) as archive:
    styles = ET.fromstring(archive.read("xl/styles.xml"))
    codes = {int(f.attrib["numFmtId"]): f.attrib["formatCode"] for f in styles.find("s:numFmts", ns)}
    xfs = list(styles.find("s:cellXfs", ns))
    sheets = ET.fromstring(archive.read("xl/workbook.xml")).find("s:sheets", ns)
    relations = {r.attrib["Id"]: r.attrib["Target"]
                 for r in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
    relationship_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    if len(sheets) != len(document.sheets):
        raise RuntimeError("The native export must have one worksheet per source sheet, without a summary.")
    for sheet, native_sheet in zip(document.sheets, sheets):
        if native_sheet.attrib["name"] != sheet.name or len(sheet.tables) != 1:
            raise RuntimeError("This reference requires matching sheet names and one source table per sheet.")
        table = sheet.tables[0]
        target = relations[native_sheet.attrib[relationship_id]]
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        xml = ET.fromstring(archive.read(path))
        cells = []
        for row in range(table.num_rows):
            for column in range(table.num_cols):
                buffer = document._model.storage_buffer(table._table_id, row, column)
                if not buffer or buffer[0] != 5 or buffer[1] != 7:
                    continue
                flags = int.from_bytes(buffer[8:12], "little")
                if not flags & (1 << 16):
                    raise RuntimeError("A selected duration cell has no explicit duration format.")
                offset = 12 + sum(16 if bit == 0 else 8 if bit in (1, 2) else 4
                                  for bit in range(16) if flags & (1 << bit))
                if offset + 4 > len(buffer):
                    raise RuntimeError("Truncated selected source duration format.")
                key = int.from_bytes(buffer[offset:offset + 4], "little")
                fmt = document._model.table_format(table._table_id, key)
                cell = table.cell(row, column)
                letters, cursor = "", column + 1
                while cursor:
                    cursor, remainder = divmod(cursor - 1, 26)
                    letters = chr(65 + remainder) + letters
                address = letters + str(row + 1)
                oracle = xml.find(".//s:c[@r='" + address + "']", ns)
                code = codes[int(xfs[int(oracle.attrib["s"])].attrib["numFmtId"])]
                serial = float(oracle.find("s:v", ns).text)
                if not math.isclose(serial * 86400, cell._double, rel_tol=1e-12, abs_tol=1e-8):
                    raise RuntimeError("The native XLSX cache differs from the independent source duration cache.")
                formula = oracle.find("s:f", ns)
                cells.append({"row": row + 1, "column": column + 1, "address": address,
                              "seconds": cell._double, "isFormula": cell.is_formula,
                              "formula": cell.formula if cell.is_formula else None,
                              "durationFormatHex": fmt.SerializeToString().hex(),
                              "settings": {f.name: v for f, v in fmt.ListFields()},
                              "nativeValueDays": serial, "nativeFormatCode": code,
                              "nativeFormula": formula.text if formula is not None else None})
        profiles.append({"sheet": sheet.name, "table": table.name,
                         "rows": table.num_rows, "columns": table.num_cols, "cells": cells})

fields = {"format_type", "duration_style", "duration_unit_largest",
          "duration_unit_smallest", "use_automatic_duration_units"}
result = {
    "extractorVersion": "numbers-parser 4.19.0",
    "qualification": "Selected fixed abbreviated duration declarations, raw seconds and typed formula caches paired with native XLSX export. Native PDF and destination application display require separate evidence.",
    "source": args.source.name,
    "sourceSha256": hashlib.sha256(args.source.read_bytes()).hexdigest(),
    "nativeExport": args.native_export.name,
    "nativeExportSha256": hashlib.sha256(args.native_export.read_bytes()).hexdigest(),
    "fieldNumbers": {f.name: f.number for f in TSKArchives_pb2.FormatStructArchive.DESCRIPTOR.fields if f.name in fields},
    "profiles": profiles,
}
args.output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
print(f"Decoded {sum(len(p['cells']) for p in profiles)} duration cells across {len(profiles)} profiles.")
