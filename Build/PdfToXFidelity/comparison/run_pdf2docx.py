"""PDF -> DOCX with pdf2docx. Usage: run_pdf2docx.py <jobs.json> <out-dir> | --one <in.pdf> <out.docx>"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if sys.argv[1] == "--one":
    from pdf2docx import Converter

    converter = Converter(sys.argv[2])
    try:
        converter.convert(sys.argv[3])
    finally:
        converter.close()
else:
    from _batch import run_batch

    run_batch(os.path.abspath(__file__), sys.argv[1], sys.argv[2], "docx")
