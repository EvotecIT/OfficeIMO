"""Inspect the root quick start's saved native documents and PDF signatures."""

from __future__ import annotations

import argparse
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from verify import RELATIONSHIP_ID, related_part


def verify(directory: Path) -> None:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    with ZipFile(directory / "Sales.docx") as archive:
        root = ET.fromstring(archive.read(related_part(archive, "", None, "officeDocument")))
        text = {node.text for node in root.findall(".//w:t", ns)}
        if not {"Quarterly sales", "North", "South", "West"} <= text:
            raise ValueError("The Word quick start is missing its heading or sales rows")
        if root.find(".//w:tbl/w:tr/w:trPr/w:tblHeader", ns) is None:
            raise ValueError("The Word sales table is missing its repeating header")

    with ZipFile(directory / "Sales.xlsx") as archive:
        workbook_path = related_part(archive, "", None, "officeDocument")
        root = ET.fromstring(archive.read(workbook_path))
        sheet = root.find("s:sheets/s:sheet[@name='Sales']", ns)
        if sheet is None:
            raise ValueError("The Excel quick start is missing its Sales sheet")
        sheet_path = related_part(archive, workbook_path, sheet.get(RELATIONSHIP_ID), "worksheet")
        root = ET.fromstring(archive.read(sheet_path))
        cells = {cell.get("r"): cell for cell in root.findall("s:sheetData/s:row/s:c", ns)}
        shared = []
        if any(cell.get("t") == "s" for cell in cells.values()):
            strings = ET.fromstring(archive.read(related_part(archive, workbook_path, None, "sharedStrings")))
            shared = ["".join(node.text or "" for node in item.findall(".//s:t", ns)) for item in strings]
        for row, region, amount in [(2, "North", 182400), (3, "South", 141900), (4, "West", 97300)]:
            first, second = cells.get(f"A{row}"), cells.get(f"B{row}")
            if first is None or second is None:
                raise ValueError("The Excel sales rows disappeared after AddChart")
            value = first.findtext("s:v", namespaces=ns)
            if first.get("t") == "s":
                text = shared[int(value)]
            elif first.get("t") == "inlineStr":
                text = "".join(node.text or "" for node in first.findall(".//s:t", ns))
            else:
                text = value
            if text != region or float(second.findtext("s:v", namespaces=ns)) != amount:
                raise ValueError(f"Unexpected Excel sales row {row}")

    with ZipFile(directory / "Sales.pptx") as archive:
        presentation_path = related_part(archive, "", None, "officeDocument")
        root = ET.fromstring(archive.read(presentation_path))
        slides = root.findall("p:sldIdLst/p:sldId", ns)
        if len(slides) != 2:
            raise ValueError("The PowerPoint quick start must contain two slides")
        titles, charts = set(), []
        for slide in slides:
            path = related_part(archive, presentation_path, slide.get(RELATIONSHIP_ID), "slide")
            content = ET.fromstring(archive.read(path))
            titles.update(node.text for node in content.findall(".//a:t", ns))
            charts.extend(content.findall(".//c:chart", ns))
        if not {"Quarterly sales", "Revenue by region"} <= titles or len(charts) != 1:
            raise ValueError("The PowerPoint quick start is missing its titles or chart")

    for name in ["Sales.pdf", "Sales-workbook.pdf", "Sales-deck.pdf"]:
        data = (directory / name).read_bytes()
        if not data.startswith(b"%PDF-") or not data.rstrip().endswith(b"%%EOF"):
            raise ValueError(f"The quick start did not save a complete PDF: {name}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Root quick start saved its native documents and PDF outputs.")
