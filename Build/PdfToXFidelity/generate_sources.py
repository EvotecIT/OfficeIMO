"""Writes feature-focused DOCX, XLSX and PPTX sources with python-docx, openpyxl and python-pptx.

Run inside the evaluation venv: python -I generate_sources.py <output-directory>
The files exercise the structures the PDF converters must recover: heading levels, nested lists,
merged and empty table cells, multi-column sections, images, links, run styles and numbers in several cultures.
"""
from __future__ import annotations

import datetime as dt
import os
import sys

from docx import Document as WordDocument
from docx.enum.section import WD_SECTION
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt

LOREM = ("Quarterly results improved across every region while operating costs stayed within the approved budget. "
         "The team delivered the migration ahead of schedule and retired two legacy systems. "
         "Customer satisfaction rose for the third consecutive quarter, driven by faster support responses.")
WORDS = LOREM.split()


def sentence(seed: int, length: int = 40) -> str:
    return " ".join(WORDS[(seed * 7 + i) % len(WORDS)] for i in range(length)).rstrip(".,") + "."


def make_image(path: str, color: tuple[int, int, int], label: str) -> str:
    img = Image.new("RGB", (480, 260), color)
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, 460, 240], outline=(255, 255, 255), width=6)
    d.ellipse([170, 60, 310, 200], fill=(255, 255, 255))
    d.text((30, 30), label, fill=(0, 0, 0))
    img.save(path)
    return path


def add_link(paragraph, url: str, text: str) -> None:
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), "Hyperlink")
    rpr.append(style)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    link.append(run)
    paragraph._p.append(link)


def set_columns(section, count: int) -> None:
    cols = section._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        section._sectPr.append(cols)
    cols.set(qn("w:num"), str(count))
    cols.set(qn("w:space"), "540")


def styled_paragraph(doc, seed: int):
    p = doc.add_paragraph()
    p.add_run("Status: ").bold = True
    p.add_run(sentence(seed, 12) + " ")
    p.add_run("Important caveat").italic = True
    p.add_run(" applies to ")
    r = p.add_run("every bold italic figure")
    r.bold = r.italic = True
    p.add_run(". ")
    big = p.add_run("Larger note.")
    big.font.size = Pt(16)
    return p


def word_headings_lists(out: str, img: str) -> None:
    doc = WordDocument()
    doc.add_heading("Annual Operations Review", 0)
    doc.add_heading("1 Summary", 1)
    doc.add_paragraph(sentence(1, 60))
    styled_paragraph(doc, 2)
    doc.add_heading("1.1 Priorities", 2)
    for i, text in enumerate(["Reduce support backlog", "Retire legacy reporting", "Expand regional coverage"]):
        doc.add_paragraph(text, style="List Bullet")
        if i == 1:
            doc.add_paragraph("Archive historic exports", style="List Bullet 2")
            doc.add_paragraph("Notify downstream owners", style="List Bullet 2")
            doc.add_paragraph("Confirm retention rules", style="List Bullet 3")
    doc.add_heading("1.2 Steps", 2)
    for text in ["Collect requirements", "Agree on scope", "Deliver the pilot", "Review outcomes"]:
        doc.add_paragraph(text, style="List Number")
    doc.add_paragraph("Validate pilot metrics", style="List Number 2")
    doc.add_heading("1.2.1 Detail", 3)
    p = doc.add_paragraph("See the project page at ")
    add_link(p, "https://example.com/operations/review", "example.com operations")
    p.add_run(" and the archive at ")
    add_link(p, "https://example.org/archive", "the archive")
    p.add_run(".")
    doc.add_heading("2 Outlook", 1)
    doc.add_paragraph(sentence(5, 80))
    doc.add_picture(img, width=Inches(3))
    doc.add_paragraph("Figure 1. Regional coverage map.")
    doc.add_paragraph(sentence(6, 50))
    doc.save(out)


def word_tables(out: str) -> None:
    doc = WordDocument()
    doc.add_heading("Table Structures", 1)
    doc.add_paragraph("A plain grid with a header row.")
    t = doc.add_table(rows=5, cols=4)
    t.style = "Table Grid"
    data = [["Region", "Q1", "Q2", "Q3"], ["North", "120", "135", "150"], ["South", "98", "101", "117"],
            ["East", "143", "", "160"], ["West", "87", "92", ""]]
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            t.cell(r, c).text = value
    doc.add_paragraph("Horizontal and vertical merges with empty cells.")
    t = doc.add_table(rows=6, cols=5)
    t.style = "Table Grid"
    a = t.cell(0, 0).merge(t.cell(1, 0))
    a.text = "Product"
    b = t.cell(0, 1).merge(t.cell(0, 2))
    b.text = "First half"
    c = t.cell(0, 3).merge(t.cell(0, 4))
    c.text = "Second half"
    for col, label in zip(range(1, 5), ["Jan-Mar", "Apr-Jun", "Jul-Sep", "Oct-Dec"]):
        t.cell(1, col).text = label
    rows = [["Widgets", "1,250", "1,310", "", "1,480"], ["Gadgets", "860", "", "910", "955"],
            ["Tools", "", "", "", ""], ["Total", "2,110", "1,310", "910", "2,435"]]
    for r, row in enumerate(rows, start=2):
        for col, value in enumerate(row):
            t.cell(r, col).text = value
    d = t.cell(2, 0).merge(t.cell(3, 0))
    d.text = "Hardware lines"
    doc.add_paragraph("A borderless layout table.")
    t = doc.add_table(rows=3, cols=3)
    for r in range(3):
        for col in range(3):
            t.cell(r, col).text = f"Item {r + 1}.{col + 1}"
    doc.add_paragraph("A long table that continues across the page break.")
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    for col, label in enumerate(["Id", "Description", "Amount"]):
        t.cell(0, col).text = label
    for i in range(1, 46):
        row = t.add_row().cells
        row[0].text = str(1000 + i)
        row[1].text = f"Line item {i} for {WORDS[i % len(WORDS)].strip('.,')}"
        row[2].text = f"{i * 37.5:,.2f}"
    doc.save(out)


def word_columns(out: str, columns: int) -> None:
    doc = WordDocument()
    doc.add_heading(f"Newsletter in {columns} columns", 1)
    doc.add_paragraph("This introduction spans the full page width before the column section starts.")
    section = doc.add_section(WD_SECTION.CONTINUOUS)
    set_columns(section, columns)
    for i in range(1, 7):
        doc.add_heading(f"Story {i}", 2)
        doc.add_paragraph(f"Story {i} begins here. " + sentence(i, 70) + f" Story {i} ends here.")
    end = doc.add_section(WD_SECTION.CONTINUOUS)
    set_columns(end, 1)
    doc.add_paragraph("This closing paragraph returns to a single column.")
    doc.save(out)


def word_cultures(out: str) -> None:
    doc = WordDocument()
    doc.add_heading("Numbers in Several Cultures", 1)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    for c, label in enumerate(["Culture", "Amount", "Percent", "Date", "Negative"]):
        t.cell(0, c).text = label
    rows = [["en-US", "1,234,567.89", "12.5%", "3/15/2024", "(1,250.00)"],
            ["de-DE", "1.234.567,89", "12,5 %", "15.03.2024", "-1.250,00"],
            ["fr-FR", "1 234 567,89", "12,5 %", "15/03/2024", "-1 250,00"],
            ["pl-PL", "1 234 567,89 zł", "12,5%", "15.03.2024", "-1 250,00 zł"],
            ["de-CH", "1'234'567.89", "12.5%", "15.03.2024", "-1'250.00"],
            ["ISO", "1234567.89", "0.125", "2024-03-15", "-1250"]]
    for row in rows:
        cells = t.add_row().cells
        for c, value in enumerate(row):
            cells[c].text = value
    doc.add_paragraph("Totals were €1.234,50 in Berlin, $1,234.50 in Boston and 1 234,50 zł in Kraków.")
    doc.save(out)


def word_mixed(out: str, img_a: str, img_b: str) -> None:
    doc = WordDocument()
    doc.add_heading("Incident Report", 0)
    doc.add_heading("Timeline", 1)
    for text in ["Alert raised at 09:12", "Engineer paged at 09:14", "Mitigation applied at 09:40"]:
        doc.add_paragraph(text, style="List Number")
    doc.add_heading("Impact", 1)
    styled_paragraph(doc, 3)
    doc.add_picture(img_a, width=Inches(2.5))
    doc.add_paragraph(sentence(8, 40))
    t = doc.add_table(rows=3, cols=3)
    t.style = "Table Grid"
    for r, row in enumerate([["Service", "Downtime", "Users"], ["API", "28 min", "4,100"], ["Portal", "12 min", "950"]]):
        for c, v in enumerate(row):
            t.cell(r, c).text = v
    doc.add_heading("Follow-up", 1)
    for text in ["Add alert on queue depth", "Document the failover", "Rehearse quarterly"]:
        doc.add_paragraph(text, style="List Bullet")
    doc.add_picture(img_b, width=Inches(2.5))
    doc.add_paragraph(sentence(9, 30))
    doc.save(out)


def excel_cultures(out: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cultures"
    ws.append(["Culture", "Amount", "Currency", "Percent", "Date", "Negative", "Count", "Flag"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    formats = [("en-US", "[$-409]#,##0.00", "[$$-409]#,##0.00", "[$-409]M/D/YYYY"),
               ("de-DE", "[$-407]#,##0.00", "#,##0.00 [$€-407]", "[$-407]DD.MM.YYYY"),
               ("fr-FR", "[$-40C]#,##0.00", "#,##0.00 [$€-40C]", "[$-40C]DD/MM/YYYY"),
               ("pl-PL", "[$-415]#,##0.00", "#,##0.00 [$zł-415]", "[$-415]DD.MM.YYYY"),
               ("ja-JP", "[$-411]#,##0", "[$¥-411]#,##0", "[$-411]YYYY/MM/DD")]
    for i, (name, number, currency, date) in enumerate(formats):
        row = 2 + i
        values = [name, 1234567.891 + i, 9876.5 * (i + 1), 0.125 + i / 100, dt.date(2024, 3, 15 + i), -1250.25 * (i + 1), 42 + i, i % 2 == 0]
        ws.append(values)
        ws.cell(row, 2).number_format = number
        ws.cell(row, 3).number_format = currency
        ws.cell(row, 4).number_format = "0.0%"
        ws.cell(row, 5).number_format = date
        ws.cell(row, 6).number_format = number.split("]")[0] + "]#,##0.00;(#,##0.00)" if "]" in number else "#,##0.00;(#,##0.00)"
    for col, width in zip("ABCDEFGH", [10, 16, 16, 10, 14, 14, 8, 8]):
        ws.column_dimensions[col].width = width
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    wb.save(out)


def excel_merged(out: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget"
    ws["A1"] = "Departmental Budget 2025"
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A3"] = "Department"
    ws.merge_cells("A3:A4")
    ws["B3"] = "H1"
    ws.merge_cells("B3:C3")
    ws["D3"] = "H2"
    ws.merge_cells("D3:E3")
    ws["F3"] = "Total"
    ws.merge_cells("F3:F4")
    for col, label in zip("BCDE", ["Q1", "Q2", "Q3", "Q4"]):
        ws[f"{col}4"] = label
    data = [("Engineering", 120000, 125000, None, 140000), ("Marketing", 45000, None, 52000, 50500),
            ("Support", 38000, 39000, 40000, 41000), ("Legal", None, None, None, None)]
    for r, row in enumerate(data, start=5):
        ws.cell(r, 1, row[0])
        for c, v in enumerate(row[1:], start=2):
            if v is not None:
                ws.cell(r, c, v).number_format = "#,##0"
        # Literal totals: formulas without cached values would leave the ground truth empty.
        ws.cell(r, 6, sum(v or 0 for v in row[1:])).number_format = "#,##0"
    ws.cell(9, 1, "Total").font = Font(bold=True)
    for c in range(2, 7):
        total = sum(ws.cell(r, c).value or 0 for r in range(5, 9))
        ws.cell(9, c, total).number_format = "#,##0"
    for col in "ABCDEF":
        ws.column_dimensions[col].width = 14
    wb.save(out)


def excel_multi(out: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Order", "Customer", "Placed", "Quantity", "Unit price", "Discount"])
    for i in range(1, 61):
        ws.append([f"SO-{2000 + i}", f"Customer {chr(65 + i % 26)}{i}", dt.date(2025, 1 + i % 12, 1 + i % 27), i * 3,
                   round(9.99 + i * 1.25, 2), (i % 5) / 20])
        ws.cell(i + 1, 3).number_format = "YYYY-MM-DD"
        ws.cell(i + 1, 5).number_format = "#,##0.00"
        ws.cell(i + 1, 6).number_format = "0%"
    ws2 = wb.create_sheet("Summary")
    ws2.append(["Metric", "Value"])
    ws2.append(["Orders", 60])
    ws2.append(["Average quantity", 91.5])
    ws2["B3"].number_format = "0.0"
    ws2.append(["Largest order", "SO-2060"])
    ws2.append(["Report date", dt.date(2025, 12, 31)])
    ws2["B5"].number_format = "DD MMM YYYY"
    for sheet in (ws, ws2):
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
        for col in "ABCDEF":
            sheet.column_dimensions[col].width = 16
    wb.save(out)


def slides(out: str, img: str) -> None:
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Platform Roadmap 2026"
    s.placeholders[1].text = "Engineering all-hands"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Goals"
    tf = s.placeholders[1].text_frame
    items = [("Ship the new editor", 0), ("Collaborative cursors", 1), ("Offline drafts", 1), ("Conflict review", 2),
             ("Cut page load time in half", 0), ("Retire the legacy API", 0)]
    tf.text = items[0][0]
    for text, level in items[1:]:
        p = tf.add_paragraph()
        p.text = text
        p.level = level
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Milestones"
    table = s.shapes.add_table(5, 4, PInches(0.5), PInches(1.6), PInches(9), PInches(3)).table
    data = [["Milestone", "Owner", "Start", "Status"], ["Editor beta", "Team A", "Feb", "On track"],
            ["Sync engine", "Team B", "Mar", ""], ["API retirement", "Team C", "May", "At risk"], ["Launch", "All", "Sep", ""]]
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            table.cell(r, c).text = v
    table.cell(4, 1).merge(table.cell(4, 2))
    table.cell(4, 1).text = "All teams, September"
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Architecture"
    s.shapes.add_picture(img, PInches(1), PInches(1.8), width=PInches(4))
    box = s.shapes.add_textbox(PInches(5.4), PInches(2), PInches(4), PInches(2)).text_frame
    box.text = "Services talk through one gateway."
    p = box.add_paragraph()
    p.text = "Each region keeps its own cache."
    r = p.runs[0]
    r.font.bold = True
    r.font.size = PPt(20)
    s = prs.slides.add_slide(prs.slide_layouts[3])
    s.shapes.title.text = "Trade-offs"
    s.placeholders[1].text = "Faster delivery"
    s.placeholders[1].text_frame.add_paragraph().text = "Smaller teams"
    s.placeholders[2].text = "Higher coordination cost"
    s.placeholders[2].text_frame.add_paragraph().text = "More integration tests"
    prs.save(out)


def main() -> None:
    out = os.path.abspath(sys.argv[1])
    os.makedirs(out, exist_ok=True)
    images = os.path.join(out, "_images")
    os.makedirs(images, exist_ok=True)
    a = make_image(os.path.join(images, "a.png"), (40, 90, 160), "Coverage")
    b = make_image(os.path.join(images, "b.png"), (170, 70, 40), "Latency")
    word_headings_lists(os.path.join(out, "gen-headings-lists.docx"), a)
    word_tables(os.path.join(out, "gen-tables.docx"))
    word_columns(os.path.join(out, "gen-two-columns.docx"), 2)
    word_columns(os.path.join(out, "gen-three-columns.docx"), 3)
    word_cultures(os.path.join(out, "gen-number-cultures.docx"))
    word_mixed(os.path.join(out, "gen-mixed-report.docx"), a, b)
    excel_cultures(os.path.join(out, "gen-number-cultures.xlsx"))
    excel_merged(os.path.join(out, "gen-merged-budget.xlsx"))
    excel_multi(os.path.join(out, "gen-orders-two-sheets.xlsx"))
    slides(os.path.join(out, "gen-roadmap.pptx"), a)


if __name__ == "__main__":
    main()
