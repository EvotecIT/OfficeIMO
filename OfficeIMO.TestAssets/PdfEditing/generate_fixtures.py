"""Opt-in independent PDF editing fixtures; ReportLab and pypdf are test tools only."""
from pathlib import Path
import hashlib
import json
import reportlab
import pypdf
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

root = Path(__file__).resolve().parent
pdfmetrics.registerFont(TTFont("Baseline", str(root.parent / "Fonts" / "OfficeIMOBaselineSans-Regular.ttf")))
source = root / "source-font-reportlab.pdf"
document = canvas.Canvas(str(source), pagesize=(595, 842), invariant=1, pageCompression=0)
document.setFont("Baseline", 14)
document.drawString(50, 700, "alpha beta gamma Żółć")
document.save()

for algorithm in ("RC4-40", "AES-128", "AES-256"):
    writer = pypdf.PdfWriter(clone_from=source)
    writer.encrypt("open", "owner", algorithm=algorithm)
    with (root / f"source-font-{algorithm.lower()}.pdf").open("wb") as output:
        writer.write(output)

manifest = {
    "producer": {"reportlab": reportlab.Version, "pypdf": pypdf.__version__},
    "text": "alpha beta gamma Żółć",
    "passwords": {"user": "open", "owner": "owner"},
    "files": {pdf.name: {"bytes": pdf.stat().st_size, "sha256": hashlib.sha256(pdf.read_bytes()).hexdigest()}
              for pdf in sorted(root.glob("*.pdf"))},
}
(root / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
