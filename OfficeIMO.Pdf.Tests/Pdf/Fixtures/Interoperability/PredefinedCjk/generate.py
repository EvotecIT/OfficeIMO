"""Opt-in independent producer; requires ReportLab 4.4.9 and pypdf 6.10.0."""
import hashlib
import io
import json
from pathlib import Path

import pypdf
import reportlab
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

root = Path(__file__).resolve().parent
cases = []
for name, face, term, encoding in [
    ("japanese", "HeiseiMin-W3", "機密", "UniJIS-UCS2-H"),
    ("simplified-chinese", "STSong-Light", "秘密", "UniGB-UCS2-H"),
    ("traditional-chinese", "MSung-Light", "秘密", "UniCNS-UCS2-H"),
    ("korean", "HYSMyeongJo-Medium", "비밀", "UniKS-UCS2-H"),
]:
    pdfmetrics.registerFont(UnicodeCIDFont(face))
    buffer = io.BytesIO()
    document = canvas.Canvas(buffer, pagesize=(612, 792), pageCompression=0, invariant=1)
    document.setFont(face, 18)
    line = f"Before {term} after page"
    document.drawString(72, 700, line)
    document.setFont("Helvetica", 12)
    document.drawString(72, 660, "Public summary remains readable.")
    document.save()
    reader = PdfReader(buffer)
    fonts = reader.pages[0]["/Resources"]["/Font"].get_object()
    for reference in fonts.values():
        font = reference.get_object()
        if font.get("/Subtype") == "/Type0":
            if name == "traditional-chinese":
                # ReportLab 4.4.9 declares UniGB with Adobe/CNS1 for MSung-Light.
                wrong = PdfWriter()
                wrong.append(reader)
                wrong.write(root / "mismatched-collection.pdf")
            font[NameObject("/Encoding")] = NameObject("/" + encoding)
    writer = PdfWriter()
    writer.append(reader)
    path = root / (name + ".pdf")
    writer.write(path)
    payload = path.read_bytes()
    cases.append({"file": path.name, "font": face, "encoding": encoding, "term": term,
                  "line": line, "advancePoints": pdfmetrics.stringWidth(line, face, 18),
                  "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()})

wrong = (root / "mismatched-collection.pdf").read_bytes()
metadata = {"producerVersions": {"ReportLab": reportlab.Version, "pypdf": pypdf.__version__},
            "documentLicense": "MIT", "runtime": "No embedded font programs or ToUnicode; predefined Adobe ROS resources",
            "traditionalChineseCorrection": "pypdf sets UniCNS-UCS2-H instead of the ReportLab 4.4.9 UniGB-UCS2-H default for Adobe/CNS1",
            "cases": cases,
            "refusalCase": {"file": "mismatched-collection.pdf", "bytes": len(wrong),
                            "sha256": hashlib.sha256(wrong).hexdigest(), "reason": "UniGB-UCS2-H with Adobe/CNS1"}}
(root / "source.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
