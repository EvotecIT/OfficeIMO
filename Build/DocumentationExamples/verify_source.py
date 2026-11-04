"""Check the saved 3.6 quick starts through their native package relationships."""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from verify import RELATIONSHIP_ID, related_part


def _worksheet_values(workbook: ZipFile, workbook_path: str, sheet_root: ET.Element, ns: dict) -> dict:
    """Read cell values using the workbook's related shared strings when needed."""
    cells = sheet_root.findall("s:sheetData/s:row/s:c", ns)
    shared = []
    if any(cell.get("t") == "s" for cell in cells):
        shared_root = ET.fromstring(workbook.read(related_part(workbook, workbook_path, None, "sharedStrings")))
        shared = ["".join(node.text or "" for node in item.findall(".//s:t", ns)) for item in shared_root]
    values = {}
    for cell in cells:
        value = cell.findtext("s:v", default="", namespaces=ns)
        if cell.get("t") == "s":
            value = shared[int(value)]
        elif cell.get("t") == "inlineStr":
            value = "".join(node.text or "" for node in cell.findall(".//s:t", ns))
        values[cell.get("r")] = value
    return values


def verify(directory: Path) -> None:
    word_ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(directory / "Status.docx") as document:
        root = ET.fromstring(document.read(related_part(document, "", None, "officeDocument")))
        text = [node.text for node in root.findall(".//w:t", word_ns)]
        if "Northwind migration" not in text or "On track" not in text:
            raise ValueError("The Word quick start is missing its status paragraphs")
        table = root.find(".//w:tbl", word_ns)
        if table is None:
            raise ValueError("The Word quick start is missing its typed table")
        rows = [["".join(node.text or "" for node in cell.findall(".//w:t", word_ns))
                 for cell in row.findall("w:tc", word_ns)] for row in table.findall("w:tr", word_ns)]
        if rows != [["Name", "Owner", "Status"], ["Discovery", "Product", "Done"],
                    ["Pilot", "Operations", "At risk"], ["Rollout", "IT", "Planned"]]:
            raise ValueError(f"Unexpected Word milestone rows: {rows}")
        header = table.find("w:tr/w:trPr/w:tblHeader", word_ns)
        if header is None or header.get("{" + word_ns["w"] + "}val", "true").lower() not in {"true", "1", "on"}:
            raise ValueError("The typed Word table does not repeat its header")

    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    with ZipFile(directory / "Orders.xlsx") as workbook:
        workbook_path = related_part(workbook, "", None, "officeDocument")
        root = ET.fromstring(workbook.read(workbook_path))
        sheet = root.find("s:sheets/s:sheet[@name='Orders']", ns)
        if sheet is None:
            raise ValueError("The Excel quick start is missing the Orders sheet")
        sheet_path = related_part(workbook, workbook_path, sheet.get(RELATIONSHIP_ID), "worksheet")
        sheet_root = ET.fromstring(workbook.read(sheet_path))
        values = _worksheet_values(workbook, workbook_path, sheet_root, ns)
        parsed_sheets = {sheet_path: values}
        expected = [["Customer", "Region", "Total", "Margin"], ["Contoso", "North", 12400, .21],
                    ["Fabrikam", "South", 8900, .17], ["Northwind", "North", 15100, .24],
                    ["Tailspin", "West", 6300, .12]]
        for row_index, row in enumerate(expected, 1):
            for column, value in zip("ABCD", row):
                address = f"{column}{row_index}"
                actual = values.get(address)
                matches = actual is not None and (
                    math.isclose(float(actual), value, abs_tol=1e-10)
                    if isinstance(value, (float, int)) else actual == value)
                if not matches:
                    raise ValueError(f"Excel table data changed after AddChart at {address}: {actual!r}")
        tables = [ET.fromstring(workbook.read(related_part(workbook, sheet_path, part.get(RELATIONSHIP_ID), "table")))
                  for part in sheet_root.findall("s:tableParts/s:tablePart", ns)]
        if not any(table.get("name") == "Orders" and table.get("ref") == "A1:D5" for table in tables):
            raise ValueError("The Excel quick start is missing its named Orders table")
        drawing = sheet_root.find("s:drawing", ns)
        if drawing is None:
            raise ValueError("The Excel quick start is missing its chart drawing")
        drawing_path = related_part(workbook, sheet_path, drawing.get(RELATIONSHIP_ID), "drawing")
        drawing_root = ET.fromstring(workbook.read(drawing_path))
        charts = drawing_root.findall(".//c:chart", ns)
        if len(charts) != 1:
            raise ValueError("The Excel quick start must contain one chart")
        chart_path = related_part(workbook, drawing_path, charts[0].get(RELATIONSHIP_ID), "chart")
        chart = ET.fromstring(workbook.read(chart_path))

        def chart_values(kind: str) -> list:
            """Keep present caches authoritative; resolve cacheless A1 ranges through workbook relationships."""
            result = []
            for data in chart.findall(f".//c:{kind}", ns):
                if any(data.find(f".//c:{tag}", ns) is not None
                       for tag in ("strCache", "numCache", "multiLvlStrCache", "strLit", "numLit")):
                    result.extend(node.text for node in data.findall(".//c:pt/c:v", ns))
                    continue
                formula = data.findtext(".//c:f", default="", namespaces=ns)
                sheet_name, separator, address_range = formula.rpartition("!")
                if sheet_name.startswith("'") and sheet_name.endswith("'"):
                    sheet_name = sheet_name[1:-1].replace("''", "'")
                referenced_sheet = next((item for item in root.findall("s:sheets/s:sheet", ns)
                                         if item.get("name", "").casefold() == sheet_name.casefold()), None)
                if not separator or referenced_sheet is None:
                    raise ValueError(f"Unresolved chart worksheet reference: {formula}")
                referenced_path = related_part(workbook, workbook_path, referenced_sheet.get(RELATIONSHIP_ID), "worksheet")
                if referenced_path not in parsed_sheets:
                    referenced_root = ET.fromstring(workbook.read(referenced_path))
                    parsed_sheets[referenced_path] = _worksheet_values(workbook, workbook_path, referenced_root, ns)
                bounds = []
                for address in address_range.split(":"):
                    match = re.fullmatch(r"\$?([A-Z]+)\$?([1-9][0-9]*)", address.upper())
                    if match is None:
                        raise ValueError(f"Unsupported chart cell reference: {formula}")
                    column = 0
                    for letter in match[1]:
                        column = column * 26 + ord(letter) - ord("A") + 1
                    bounds.append((int(match[2]), column))
                if len(bounds) not in (1, 2) or bounds[0][0] > bounds[-1][0] or bounds[0][1] > bounds[-1][1]:
                    raise ValueError(f"Invalid chart cell range: {formula}")
                for row in range(bounds[0][0], bounds[-1][0] + 1):
                    for column in range(bounds[0][1], bounds[-1][1] + 1):
                        letters = ""
                        number = column
                        while number:
                            number, remainder = divmod(number - 1, 26)
                            letters = chr(ord("A") + remainder) + letters
                        result.append(parsed_sheets[referenced_path][f"{letters}{row}"])
            return result

        categories = chart_values("cat")
        amounts = [float(value) for value in chart_values("val")]
        if categories != ["Contoso", "Fabrikam", "Northwind", "Tailspin"] or amounts != [12400, 8900, 15100, 6300]:
            raise ValueError(f"Unexpected Excel chart data: {categories}, {amounts}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Current-source Word and Excel quick starts passed native package checks.")
