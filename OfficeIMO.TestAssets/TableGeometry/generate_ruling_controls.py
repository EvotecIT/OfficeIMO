"""Generate independent closed-ruling admission controls with ReportLab 4.4.9.

Run from an isolated contributor environment with the pinned ReportLab version.
Only this family's PDFs and ruling-controls-manifest.json are written. The original
generator supplies authored constants through runpy without executing its main.
"""

import hashlib
import json
from pathlib import Path
import runpy

import reportlab
from reportlab.pdfgen.canvas import Canvas


ROOT = Path(__file__).resolve().parent
AUTHORING = runpy.run_path(str(ROOT / "generate.py"))
REPORTLAB_VERSION = AUTHORING["REPORTLAB_VERSION"]
PAGE_SIZE = AUTHORING["PAGE_SIZE"]
COLUMNS = AUTHORING["COLUMN_BOUNDARIES"]
ROWS = AUTHORING["ROW_BOUNDARIES"]
TABLE = AUTHORING["TABLE"]
IDENTITY = (1, 0, 0, 1, 0, 0)
QUARTER_TURN = (0, -1, 1, 0, 0, 612)
CONTROLS = [
    (
        "no-rules",
        False,
        "Aligned text alone does not establish a closed ruling grid.",
        ["Omit every ruling; preserve the ledger text and placement."],
    ),
    (
        "open-right",
        False,
        "The missing right outer edge leaves the grid open.",
        ["Omit the complete vertical ruling at x=570 points."],
    ),
    (
        "missing-row-edge",
        False,
        "An interior row edge has an unresolved gap within a cell.",
        ["At y=630 points, omit the horizontal segment from x=354 to x=450."],
    ),
    (
        "ambiguous-column",
        False,
        "An extra full-height boundary creates an unpopulated ambiguous column.",
        ["Add a vertical ruling at x=250 from y=560 to y=700 in the first cell's empty margin."],
    ),
    (
        "decorative-lines",
        True,
        "Disconnected decoration outside the complete grid does not change its cells.",
        [
            "Add a horizontal line from (42,500) to (570,500).",
            "Add a vertical line from (20,530) to (20,740).",
            "Add a diagonal line from (42,480) to (570,450).",
        ],
    ),
    (
        "clipped-grid",
        False,
        "Every ruling is clipped away; invisible operations cannot establish a visible grid.",
        ["Clip all grid strokes to rectangle [0,0,10,10]; restore state before drawing text."],
    ),
    (
        "dashed-grid",
        False,
        "Dashed rulings are outside the continuous closed-grid admission profile.",
        ["Use a 4-point dash and 3-point gap on every ruling; restore state before text."],
    ),
    (
        "content-quarter-turn",
        True,
        "A complete closed grid remains admissible after an explicit content-matrix quarter-turn.",
        [
            "Use landscape MediaBox [0,0,792,612] with /Rotate=0.",
            "Apply content matrix [0,-1,1,0,0,612] to grid and text together.",
        ],
    ),
    (
        "zero-stroke-opacity",
        False,
        "Fully transparent strokes do not establish a visible ruling grid.",
        ["Set stroke opacity to zero for grid operations; restore state before text."],
    ),
    (
        "partial-clip",
        False,
        "A rectangular clip removes the right outer ruling and truncates its incident edges.",
        ["Clip grid strokes to rectangle [40,558,568,702]; restore state before text."],
    ),
    (
        "containing-clip",
        True,
        "An exact rectangular clip that contains every stroke preserves the complete grid.",
        ["Clip grid strokes to rectangle [40,558,572,702]; restore state before text."],
    ),
    (
        "opaque-cover",
        False,
        "Later opaque paint removes part of a ruling from the visible grid.",
        ["After the grid and before text, cover rectangle [150,628,180,632] with opaque white fill."],
    ),
    (
        "multiply-blend",
        False,
        "Nonnormal graphics effects are outside the ordinary-stroke admission profile.",
        ["Set blend mode Multiply for grid operations; restore state before text."],
    ),
    (
        "assignment-conflict",
        False,
        "A complete grid cannot refine the ledger when a moved separator crosses Units source text.",
        ["Replace the vertical boundary at x=282 with x=342; preserve text anchors and horizontal edges."],
    ),
]


def draw_grid(canvas, name):
    if name == "no-rules":
        return

    canvas.saveState()
    clip_rectangle = {
        "clipped-grid": (0, 0, 10, 10),
        "partial-clip": (40, 558, 568, 702),
        "containing-clip": (40, 558, 572, 702),
    }.get(name)
    if clip_rectangle is not None:
        left, bottom, right, top = clip_rectangle
        clipping = canvas.beginPath()
        clipping.rect(left, bottom, right - left, top - bottom)
        canvas.clipPath(clipping, stroke=0, fill=0)
    if name == "dashed-grid":
        canvas.setDash(4, 3)
    if name == "zero-stroke-opacity":
        canvas.setStrokeAlpha(0)
    if name == "multiply-blend":
        canvas.setBlendMode("Multiply")

    for x in COLUMNS:
        if name == "open-right" and x == COLUMNS[-1]:
            continue
        actual_x = 342 if name == "assignment-conflict" and x == 282 else x
        canvas.line(actual_x, ROWS[-1], actual_x, ROWS[0])
    for y in ROWS:
        if name == "missing-row-edge" and y == 630:
            canvas.line(COLUMNS[0], y, 354, y)
            canvas.line(450, y, COLUMNS[-1], y)
        else:
            canvas.line(COLUMNS[0], y, COLUMNS[-1], y)
    if name == "ambiguous-column":
        canvas.line(250, ROWS[-1], 250, ROWS[0])
    canvas.restoreState()

    if name == "decorative-lines":
        canvas.line(42, 500, 570, 500)
        canvas.line(20, 530, 20, 740)
        canvas.line(42, 480, 570, 450)
    if name == "opaque-cover":
        canvas.saveState()
        canvas.setFillColorRGB(1, 1, 1)
        canvas.rect(150, 628, 30, 4, stroke=0, fill=1)
        canvas.restoreState()


def draw_text(canvas):
    canvas.setFont(AUTHORING["FONT"], AUTHORING["FONT_SIZE"])
    for row_index, row in enumerate(TABLE):
        y = ROWS[row_index] - AUTHORING["BASELINE_OFFSET"]
        for column_index, text in enumerate(row):
            if AUTHORING["ALIGNMENTS"][column_index] == "left":
                canvas.drawString(COLUMNS[column_index] + AUTHORING["CELL_INSET"], y, text)
            else:
                canvas.drawRightString(
                    COLUMNS[column_index + 1] - AUTHORING["CELL_INSET"], y, text
                )


def grid_geometry(matrix, page_height):
    widths = [right - left for left, right in zip(COLUMNS, COLUMNS[1:])]
    strips = []
    a, b, c, d, e, f = matrix
    for left, right in zip(COLUMNS, COLUMNS[1:]):
        corners = [
            (a * x + c * y + e, page_height - (b * x + d * y + f))
            for x, y in (
                (left, ROWS[-1]),
                (right, ROWS[-1]),
                (right, ROWS[0]),
                (left, ROWS[0]),
            )
        ]
        strips.append([
            min(x for x, _ in corners),
            min(y for _, y in corners),
            max(x for x, _ in corners),
            max(y for _, y in corners),
        ])
    return {
        "coordinateSystem": "Ledger-local coordinates before contentMatrix",
        "columnBoundariesPoints": list(COLUMNS),
        "columnWidthsPoints": widths,
        "normalizedColumnWidths": [width / sum(widths) for width in widths],
        "rowBoundariesPoints": list(ROWS),
        "boundaryReference": "Ruling stroke centre",
        "columnStripsVisualPoints": strips,
        "visualCoordinateSystem": "Page top-left; [left,top,right,bottom] in points",
    }


def generate_case(name, expected_admission, reason, changes):
    quarter_turn = name == "content-quarter-turn"
    page_size = (792, 612) if quarter_turn else PAGE_SIZE
    matrix = QUARTER_TURN if quarter_turn else IDENTITY
    case_id = f"ruled-control-{name}"
    path = ROOT / f"{case_id}-native.pdf"
    canvas = Canvas(
        str(path),
        pagesize=page_size,
        invariant=True,
        pageCompression=0,
        pdfVersion=(1, 4),
    )
    canvas.setTitle(f"Ruled ledger admission control: {name}")
    canvas.setAuthor("OfficeIMO fixture authors")
    canvas.setSubject("Independent closed-grid evidence and conservative admission controls")
    canvas.setPageRotation(0)
    canvas.transform(*matrix)
    canvas.setLineWidth(AUTHORING["RULING_WIDTH"])
    canvas.setStrokeColorRGB(0, 0, 0)
    canvas.setFillColorRGB(0, 0, 0)
    draw_grid(canvas, name)
    draw_text(canvas)
    canvas.showPage()
    canvas.save()

    result = {
        "id": case_id,
        "expectedAdmission": expected_admission,
        "expectationReason": reason,
        "pageSizePoints": list(page_size),
        "clockwiseDegrees": 0,
        "table": TABLE,
        "readingOrder": [" ".join(row) for row in TABLE],
        "construction": {
            "contentMatrix": list(matrix),
            "changes": changes,
            "rulingWidthPoints": AUTHORING["RULING_WIDTH"],
            "font": AUTHORING["FONT"],
            "fontSizePoints": AUTHORING["FONT_SIZE"],
            "fontEmbedded": False,
            "textBaselineOffsetPoints": AUTHORING["BASELINE_OFFSET"],
            "cellInsetPoints": AUTHORING["CELL_INSET"],
            "alignments": list(AUTHORING["ALIGNMENTS"]),
        },
        "files": {"native.pdf": hashlib.sha256(path.read_bytes()).hexdigest()},
    }
    if expected_admission:
        result["geometry"] = grid_geometry(matrix, page_size[1])
    clip_rectangle = {
        "clipped-grid": [0, 0, 10, 10],
        "partial-clip": [40, 558, 568, 702],
        "containing-clip": [40, 558, 572, 702],
    }.get(name)
    if clip_rectangle is not None:
        result["construction"]["clipRectanglePoints"] = clip_rectangle
    if name == "zero-stroke-opacity":
        result["construction"]["strokeOpacity"] = 0
    if name == "multiply-blend":
        result["construction"]["blendMode"] = "Multiply"
    if name == "opaque-cover":
        result["construction"]["opaqueCover"] = {
            "boundsPoints": [150, 628, 180, 632],
            "fillRgb": [1, 1, 1],
            "afterGrid": True,
            "beforeText": True,
        }
    if name == "assignment-conflict":
        result["construction"]["replacedVerticalBoundaryPoints"] = [282, 342]
    return result


def main():
    if reportlab.Version != REPORTLAB_VERSION:
        raise RuntimeError(f"Use ReportLab {REPORTLAB_VERSION}; found {reportlab.Version}.")
    manifest = {
        "schemaVersion": 1,
        "producer": "ReportLab",
        "reportlab": reportlab.Version,
        "sourceCoordinateSystem": "PDF default user space; points; origin lower-left",
        "expectationScope": "Admission as a complete continuous visible ruling grid; aligned-text fallback remains separate",
        "license": "MIT; repository LICENSE covers authored fixtures and labels",
        "producerLicense": "ReportLab BSD license; generation tooling only",
        "fontProvenance": "Unembedded PDF standard 14 Helvetica; no font binary bundled",
        "cases": [generate_case(*control) for control in CONTROLS],
    }
    (ROOT / "ruling-controls-manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8", newline="\n"
    )


if __name__ == "__main__":
    main()
