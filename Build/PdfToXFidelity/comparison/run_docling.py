"""PDF -> Markdown and HTML with Docling's standard PDF pipeline (layout and table models; OCR off for born-digital input).

Usage: run_docling.py <jobs.json> <out-dir>
"""
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _batch import write_log  # noqa: E402
from docling.datamodel.base_models import ConversionStatus, InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.document_converter import DocumentConverter, PdfFormatOption

jobs_path, out_dir = sys.argv[1], sys.argv[2]
os.makedirs(out_dir, exist_ok=True)
options = PdfPipelineOptions()
options.do_ocr = False
options.do_table_structure = True
options.document_timeout = 180
converter = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=options)})
with open(jobs_path, encoding="utf-8") as f:
    jobs = json.load(f)
log = []
write_log(out_dir, log)
for job in jobs:
    for ext in (".md", ".html"):
        if os.path.exists(os.path.join(out_dir, job["id"] + ext)):
            os.remove(os.path.join(out_dir, job["id"] + ext))
    start = time.perf_counter()
    try:
        result = converter.convert(job["pdf"], raises_on_error=True)
        document = result.document
        with open(os.path.join(out_dir, job["id"] + ".md"), "w", encoding="utf-8") as f:
            f.write(document.export_to_markdown())
        with open(os.path.join(out_dir, job["id"] + ".html"), "w", encoding="utf-8") as f:
            f.write(document.export_to_html())
        # PARTIAL_SUCCESS keeps whatever pages converted; only a full conversion is scored as a success.
        status = "ok" if result.status == ConversionStatus.SUCCESS else str(result.status)
    except Exception as exc:  # recorded and scored as a failed conversion
        status = "error: " + type(exc).__name__ + ": " + str(exc)[:200]
    log.append({"id": job["id"], "status": status, "seconds": round(time.perf_counter() - start, 2)})
    print(job["id"], status[:80], flush=True)
    write_log(out_dir, log)
