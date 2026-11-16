"""PDF -> Markdown and HTML with Docling's standard PDF pipeline (layout and table models; OCR off for born-digital input).

Usage: run_docling.py <jobs.json> <out-dir>
"""
import json
import os
import sys
import time

from docling.datamodel.base_models import InputFormat
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
for job in jobs:
    start = time.perf_counter()
    try:
        result = converter.convert(job["pdf"], raises_on_error=True)
        document = result.document
        with open(os.path.join(out_dir, job["id"] + ".md"), "w", encoding="utf-8") as f:
            f.write(document.export_to_markdown())
        with open(os.path.join(out_dir, job["id"] + ".html"), "w", encoding="utf-8") as f:
            f.write(document.export_to_html())
        status = str(result.status)
    except Exception as exc:  # recorded and scored as a failed conversion
        status = "error: " + type(exc).__name__ + ": " + str(exc)[:200]
    log.append({"id": job["id"], "status": status, "seconds": round(time.perf_counter() - start, 2)})
    print(job["id"], status[:80], flush=True)
with open(os.path.join(out_dir, "_log.json"), "w", encoding="utf-8") as f:
    json.dump(log, f, indent=1)
