"""Inspect actual vector paint in Office's PDF exports (opt-in: requires MuPDF's mutool)."""
import json
import subprocess
import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
BLUE = (21 / 255, 96 / 255, 130 / 255)
ORANGE = (233 / 255, 113 / 255, 50 / 255)


def has_paint(pages, tag, color):
    for page in pages:
        for element in page.iter(tag):
            components = tuple(float(value) for value in element.get('color', '').split())
            if len(components) == 3 and all(abs(a - b) < .001 for a, b in zip(components, color)):
                return True
    return False


results = []
for producer in ('word', 'powerpoint', 'excel'):
    path = HERE / f'microsoft-{producer}-16.113-series-paint.pdf'
    trace = subprocess.run(['mutool', 'draw', '-F', 'trace', str(path)], capture_output=True, check=True).stdout
    pages = ET.fromstring(trace).findall('page')
    stride = 2 if producer == 'excel' else 1
    assert len(pages) == 16 * stride, (producer, len(pages))
    for index in range(16):
        selected = pages[index * stride:(index + 1) * stride]
        fill = has_paint(selected, 'fill_path', BLUE)
        outline = has_paint(selected, 'stroke_path', ORANGE)
        expected_fill = index < 12 and index % 4 < 2
        expected_outline = index % 2 == 0
        assert (fill, outline) == (expected_fill, expected_outline), (producer, index + 1, fill, outline)
    source = HERE / ('series-paint.' + {'word': 'docx', 'powerpoint': 'pptx', 'excel': 'xlsx'}[producer])
    results.append({'producer': producer, 'source': source.name, 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                    'pdf': path.name, 'pdfSha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'qualifiedPaintCombinations': 16})
print(json.dumps({'schemaVersion': 1, 'exports': results}, indent=2))
