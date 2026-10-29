"""Read the quick-start documents independently of OfficeIMO."""

import argparse
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

RELATIONSHIP_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def related_part(archive: ZipFile, part: str, relationship_id: str, kind: str) -> str:
    """Resolve one typed, internal package relationship from a generated example."""
    relationships = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    root = ET.fromstring(archive.read(relationships))
    matches = [node for node in root if node.get("Id") == relationship_id]
    expected_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind
    if len(matches) != 1 or matches[0].get("Type") != expected_type or matches[0].get("TargetMode", "Internal") != "Internal":
        raise ValueError(f"Invalid {kind} relationship from {part}")
    target = matches[0].get("Target", "")
    resolved = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(part), target))
    if not target or resolved == ".." or resolved.startswith("../") or resolved not in archive.namelist():
        raise ValueError(f"Missing {kind} relationship target from {part}")
    return resolved


def verify(directory: Path) -> None:
    """Check trusted quick-start output, including the relationships that attach its worksheet and table."""
    word_ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(directory / "report.docx") as document:
        root = ET.fromstring(document.read("word/document.xml"))
        text = [node.text for node in root.findall(".//w:t", word_ns)]
        if "Quarterly report" not in text or "Created with OfficeIMO.Word." not in text:
            raise ValueError("Word quick start is missing its report paragraphs")
        table = root.find(".//w:tbl", word_ns)
        if table is None:
            raise ValueError("Word quick start is missing its table")
        rows = [["".join(cell.itertext()) for cell in row.findall("w:tc", word_ns)]
                for row in table.findall("w:tr", word_ns)]
        if rows != [["Area", "Status"], ["Documents", "Generated"]]:
            raise ValueError(f"Unexpected Word table contents: {rows}")
        first_row = table.find("w:tr", word_ns)
        header = first_row.find("w:trPr/w:tblHeader", word_ns) if first_row is not None else None
        if header is None or header.get("{" + word_ns["w"] + "}val", "true").lower() not in {"true", "1", "on"}:
            raise ValueError("Word quick start does not repeat its header row")

    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(directory / "report.xlsx") as workbook:
        root = ET.fromstring(workbook.read("xl/workbook.xml"))
        sheet = root.find("s:sheets/s:sheet", ns)
        if sheet is None or sheet.get("name") != "Data":
            raise ValueError("Excel quick start is missing the Data worksheet")
        shared = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared = ["".join(item.itertext()) for item in shared_root]
        sheet_path = related_part(workbook, "xl/workbook.xml", sheet.get(RELATIONSHIP_ID, ""), "worksheet")
        sheet_root = ET.fromstring(workbook.read(sheet_path))
        values = {}
        for cell in sheet_root.findall("s:sheetData/s:row/s:c", ns):
            value = cell.findtext("s:v", default="", namespaces=ns)
            if cell.get("t") == "s":
                value = shared[int(value)]
            elif cell.get("t") == "inlineStr":
                value = "".join(node.text or "" for node in cell.findall(".//s:t", ns))
            values[cell.get("r")] = value
        if values != {"A1": "Name", "B1": "Value", "A2": "Alpha", "B2": "42"}:
            raise ValueError(f"Unexpected Excel cell contents: {values}")
        tables = [ET.fromstring(workbook.read(related_part(workbook, sheet_path, node.get(RELATIONSHIP_ID, ""), "table")))
                  for node in sheet_root.findall("s:tableParts/s:tablePart", ns)]
        if not any(table.get("name") == "DataTable" and table.get("ref") == "A1:B2" for table in tables):
            raise ValueError("Excel quick start is missing the named DataTable over A1:B2")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Word and Excel quick-start artifacts verified.")
