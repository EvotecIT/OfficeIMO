"""PDF -> Markdown with PyMuPDF4LLM (AGPL; evaluation only, never redistributed).

Usage: run_pymupdf4llm.py <jobs.json> <out-dir> | --one <in.pdf> <out.md>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if sys.argv[1] == "--one":
    import pymupdf4llm

    text = pymupdf4llm.to_markdown(sys.argv[2])
    with open(sys.argv[3], "w", encoding="utf-8") as f:
        f.write(text)
else:
    from _batch import run_batch

    run_batch(os.path.abspath(__file__), sys.argv[1], sys.argv[2], "md")
