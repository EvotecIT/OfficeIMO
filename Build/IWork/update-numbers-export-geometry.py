"""Record grid geometry from the pinned Numbers formula PDF (opt-in pdfplumber tool)."""
import argparse
import hashlib
import json
import re
from pathlib import Path

import pdfplumber

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("manifest", type=Path)
args = parser.parse_args()
manifest = args.manifest.resolve()
evidence = json.loads(manifest.read_text())
if evidence["sourceFixture"] != "numbers-parser/test-10-formulas.numbers":
    raise ValueError("This extractor qualifies the pinned two-table formula fixture only.")
artifacts = [a for a in evidence["artifacts"] if a["path"].endswith(".pdf")]
if len(artifacts) != 1:
    raise ValueError("Expected one pinned PDF reference.")
artifact = artifacts[0]
pdf_path = manifest.parent / artifact["path"]
if hashlib.sha256(pdf_path.read_bytes()).hexdigest() != artifact["sha256"]:
    raise ValueError("The PDF reference does not match its manifest hash.")

# Table counts come from the reference's coordinate mapping, not PDF line counts.
shapes = {}
for formula in evidence["formulaExpectations"]:
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", formula["sourceCell"])
    if not match:
        raise ValueError("Unsupported source cell coordinate.")
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord("A") + 1
    index = formula["tableIndex"]
    rows, columns = shapes.get(index, (0, 0))
    shapes[index] = (max(rows, int(match[2])), max(columns, column))
if sorted(shapes) != [1, 2]:
    raise ValueError("Expected two mapped tables.")

measurements = []
with pdfplumber.open(pdf_path) as pdf:
    if len(pdf.pages) != 2:
        raise ValueError("Expected one table on each of two reference pages.")
    for index, page in enumerate(pdf.pages, 1):
        rows, columns = shapes[index]
        vertical, horizontal = [], []
        for line in page.lines:
            x0, y0, x1, y1 = (line[k] for k in ("x0", "y0", "x1", "y1"))
            if abs(x1 - x0) < 0.001 and y1 - y0 > 1:
                vertical.append(round(x0, 3))
            elif abs(y1 - y0) < 0.001 and x1 - x0 > 1:
                horizontal.append(round(y0, 3))
            else:
                raise ValueError("Reference contains non-grid line geometry.")
        xs, ys = sorted(set(vertical)), sorted(set(horizontal), reverse=True)
        if len(xs) != columns + 1 or len(ys) != rows + 1:
            raise ValueError("Grid line counts disagree with mapped table dimensions.")
        widths = [round(b - a, 3) for a, b in zip(xs, xs[1:])]
        heights = [round(a - b, 3) for a, b in zip(ys, ys[1:])]
        if min(widths + heights) <= 0:
            raise ValueError("Grid dimensions must be positive.")
        measurements.append({"tableIndex": index, "pageNumber": index,
            "columnWidthsPoints": widths, "rowHeightsPoints": heights})

evidence["pdf"]["tableGeometry"] = {
    "measurement": "Axis-aligned stroke centre positions, rounded to 0.001 point",
    "tool": "Build/IWork/update-numbers-export-geometry.py",
    "pdfplumberVersion": pdfplumber.__version__,
    "tables": measurements,
    "limitations": "This measures the pinned rendered grid only. It does not qualify font metrics, styles, content, page placement or automatic row sizing."
}
manifest.write_text(json.dumps(evidence, indent=2) + "\n")
print(json.dumps(evidence["pdf"]["tableGeometry"], indent=2))
