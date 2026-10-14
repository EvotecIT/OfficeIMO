"""Generate independent native PDF table geometry fixtures with ReportLab 4.4.9."""

import hashlib
import json
from pathlib import Path

import reportlab
from reportlab.pdfbase.pdfdoc import PDFArray
from reportlab.pdfgen.canvas import Canvas


ROOT = Path(__file__).resolve().parent
REPORTLAB_VERSION = "4.4.9"
PAGE_SIZE = (612, 792)
COLUMN_BOUNDARIES = (42, 282, 354, 450, 570)
ROW_BOUNDARIES = (700, 665, 630, 595, 560)
FONT = "Helvetica"
FONT_SIZE = 12
CELL_INSET = 8
BASELINE_OFFSET = 22
RULING_WIDTH = 0.5
ALIGNMENTS = ("left", "right", "right", "right")
TABLE = [
    ["Description", "Units", "Rate", "Amount"],
    ["Initial service", "12", "9.99", "119.88"],
    ["Return credit", "3", "-4.50", "-13.50"],
    ["Monthly plan", "1", "89.00", "89.00"],
]


def generate_case(clockwise_degrees):
    case_id = f"ruled-ledger-{clockwise_degrees}"
    path = ROOT / f"{case_id}-native.pdf"
    canvas = Canvas(
        str(path),
        pagesize=PAGE_SIZE,
        invariant=True,
        pageCompression=0,
        pdfVersion=(1, 4),
    )
    canvas.setTitle("Unequal ruled service ledger")
    canvas.setAuthor("OfficeIMO fixture authors")
    canvas.setSubject("Independent authored column boundaries and numeric text")
    canvas.setPageRotation(clockwise_degrees)
    canvas.setLineWidth(RULING_WIDTH)
    canvas.setStrokeColorRGB(0, 0, 0)
    canvas.setFillColorRGB(0, 0, 0)

    for x in COLUMN_BOUNDARIES:
        canvas.line(x, ROW_BOUNDARIES[-1], x, ROW_BOUNDARIES[0])
    for y in ROW_BOUNDARIES:
        canvas.line(COLUMN_BOUNDARIES[0], y, COLUMN_BOUNDARIES[-1], y)

    canvas.setFont(FONT, FONT_SIZE)
    for row_index, row in enumerate(TABLE):
        y = ROW_BOUNDARIES[row_index] - BASELINE_OFFSET
        for column_index, text in enumerate(row):
            if ALIGNMENTS[column_index] == "left":
                canvas.drawString(COLUMN_BOUNDARIES[column_index] + CELL_INSET, y, text)
            else:
                canvas.drawRightString(
                    COLUMN_BOUNDARIES[column_index + 1] - CELL_INSET, y, text
                )

    canvas.showPage()
    # ReportLab otherwise swaps its default MediaBox for quarter-turn pages.
    # Fix the authored box explicitly so only /Rotate changes the displayed axes.
    canvas._doc.Pages.pages[-1].MediaBox = PDFArray([0, 0, *PAGE_SIZE])
    canvas.save()

    widths = [b - a for a, b in zip(COLUMN_BOUNDARIES, COLUMN_BOUNDARIES[1:])]
    total_width = sum(widths)
    visual_page_size = PAGE_SIZE if clockwise_degrees == 0 else tuple(reversed(PAGE_SIZE))
    return {
        "id": case_id,
        "clockwiseDegrees": clockwise_degrees,
        "pageSizePoints": list(PAGE_SIZE),
        "visualPageSizePoints": list(visual_page_size),
        "rotationMethod": "PDF /Rotate; identical content stream and MediaBox",
        "readingOrder": [" ".join(row) for row in TABLE],
        "table": TABLE,
        "geometry": {
            "columnBoundariesPoints": list(COLUMN_BOUNDARIES),
            "columnWidthsPoints": widths,
            "normalizedColumnWidths": [width / total_width for width in widths],
            "rowBoundariesPoints": list(ROW_BOUNDARIES),
            "rulingWidthPoints": RULING_WIDTH,
            "boundaryReference": "Ruling stroke centre in unrotated PDF user space",
            "font": FONT,
            "fontSizePoints": FONT_SIZE,
            "fontEmbedded": False,
            "textBaselineOffsetPoints": BASELINE_OFFSET,
            "cellInsetPoints": CELL_INSET,
            "alignments": list(ALIGNMENTS),
        },
        "acceptance": {
            "requireExactTables": True,
            "ratioReference": "Canonical detected column progression lengths",
            "maximumXlsxNormalizedShareError": 0.000001,
            "maximumOdsNormalizedShareError": 0.02,
            "sourceGridRecoveryQualified": False,
        },
        "files": {"native.pdf": hashlib.sha256(path.read_bytes()).hexdigest()},
    }


def main():
    if reportlab.Version != REPORTLAB_VERSION:
        raise RuntimeError(
            f"Use ReportLab {REPORTLAB_VERSION}; found {reportlab.Version}."
        )
    cases = [generate_case(degrees) for degrees in (0, 90)]
    manifest = {
        "schemaVersion": 1,
        "producer": "ReportLab",
        "reportlab": reportlab.Version,
        "sourceCoordinateSystem": "PDF default user space; points; origin lower-left",
        "license": "MIT; repository LICENSE covers authored fixtures and labels",
        "producerLicense": "ReportLab BSD license; generation tooling only",
        "fontProvenance": "Unembedded PDF standard 14 Helvetica; no font binary bundled",
        "cases": cases,
    }
    (ROOT / "manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8", newline="\n"
    )


if __name__ == "__main__":
    main()
