"""Read the quick-start documents independently of OfficeIMO."""

import argparse
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def verify(directory: Path) -> None:
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
        if table.find("w:tr/w:trPr/w:tblHeader", word_ns) is None:
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
        sheet_root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
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
        tables = [ET.fromstring(workbook.read(name)) for name in workbook.namelist()
                  if name.startswith("xl/tables/") and name.endswith(".xml")]
        if not any(table.get("name") == "DataTable" and table.get("ref") == "A1:B2" for table in tables):
            raise ValueError("Excel quick start is missing the named DataTable over A1:B2")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Word and Excel quick-start artifacts verified.")
