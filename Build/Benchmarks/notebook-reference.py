"""Opt-in Jupyter reference worker. PowerForge starts and measures each operation."""

import argparse
import base64
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import random
import struct
import sys
import traceback
import zlib


def make_png(pixels, seed):
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    rng = random.Random(seed)
    rows = [b"\x00" + rng.randbytes(3 * pixels) for _ in range(pixels)]
    header = struct.pack(">IIBBBBB", pixels, pixels, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b"")


def generate(path, count, code_chars, output_chars, image_pixels):
    metadata = {"language_info": {"name": "python"}}
    cells = []
    plot = base64.b64encode(make_png(image_pixels, 521)).decode("ascii") if image_pixels else ""
    attachment = base64.b64encode(make_png(image_pixels, 783)).decode("ascii") if image_pixels else ""
    for index in range(count):
        marker = f"{index:06d}"
        cell = {"id": f"cell-{marker}", "metadata": {}}
        if index % 2 == 0:
            source = [f"## NotebookCell{marker}\n", f"Stored notebook description {marker}."]
            if plot and index == 0:
                source += ["\n\n| Column | Value |\n", "| --- | --- |\n", "| TableWitness | 42 |\n"]
                source += ["\n![Attachment witness](attachment:witness.png)\n"]
                cell["attachments"] = {"witness.png": {"image/png": attachment}}
            cell.update(cell_type="markdown", source=source)
        else:
            comment = (" sample " * ((code_chars + 7) // 8))[:code_chars]
            stream = (" output " * ((output_chars + 7) // 8))[:output_chars]
            outputs = [{"output_type": "stream", "name": "stdout", "text": [f"StreamBegin{marker}", stream, f" StreamEnd{marker}\n"]}]
            if plot and index == 1:
                outputs.append({
                    "output_type": "display_data",
                    "data": {"image/png": plot, "text/plain": "Plot fallback"},
                    "metadata": {"image/png": {"width": image_pixels, "height": image_pixels}},
                })
            cell.update(cell_type="code", execution_count=index, outputs=outputs,
                        source=f"print('CodeBegin{marker}') #{comment} CodeEnd{marker}\n")
        cells.append(cell)
    notebook = {"nbformat": 4, "nbformat_minor": 5, "metadata": metadata, "cells": cells}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((json.dumps(notebook, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))


def encode_json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def notebook_sources(notebook):
    cells = notebook.get("cells")
    if cells is None:
        cells = [cell for sheet in notebook["worksheets"] for cell in sheet.get("cells", [])]
    result = []
    for cell in cells:
        kind = cell["cell_type"]
        source = "".join(cell.get("source", cell.get("input", "")))
        if kind == "heading":
            source = "#" * cell.get("level", 1) + " " + source
            kind = "markdown"
        result.append([kind, source])
    return result


class ReferenceWorker:
    def __init__(self):
        self.operation = None
        self.exporter = None
        self.artifact = None
        self.result = None
        self.resources = None

    def prepare(self, request):
        import nbformat

        self.operation = request["operation"]
        self.artifact = Path(request["artifact"])
        self.source = Path(request["input"]).read_bytes()
        self.result = None
        self.resources = None
        if self.operation in ("Markdown", "Html"):
            from nbconvert import HTMLExporter, MarkdownExporter
            from traitlets.config import Config

            config = Config({"ExecutePreprocessor": {"enabled": False}})
            self.notebook = nbformat.reads(self.source.decode("utf-8-sig"), as_version=4)
            if self.operation == "Markdown":
                self.exporter = MarkdownExporter(config=config)
            else:
                self.exporter = HTMLExporter(template_name="lab", config=config)
        elif self.operation == "Pdf":
            from nbconvert.exporters.webpdf import WebPDFExporter
            from traitlets.config import Config

            # The installed Chromium is an explicit benchmark prerequisite. Restrict
            # untrusted notebook output and assets before launching its page.
            class OfflineWebPdfExporter(WebPDFExporter):
                def run_playwright(self, html):
                    policy = (
                        '<meta http-equiv="Content-Security-Policy" '
                        'content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data:; font-src data:">'
                    )
                    return super().run_playwright(html.replace("<head>", "<head>" + policy, 1))

            self.exporter = OfflineWebPdfExporter(
                template_name="lab", allow_chromium_download=False,
                config=Config({"ExecutePreprocessor": {"enabled": False}}),
            )

    def run(self, _request):
        import nbformat

        if self.operation == "Parse":
            self.result = nbformat.reader.reads(self.source)
        elif self.operation == "SourceProjection":
            self.result = notebook_sources(nbformat.reader.reads(self.source))
        elif self.operation in ("RoundTrip", "RichRoundTrip"):
            notebook = nbformat.reads(self.source.decode("utf-8-sig"), as_version=4)
            self.result = nbformat.writes(notebook).encode("utf-8")
        elif self.operation in ("Markdown", "Html"):
            output, self.resources = self.exporter.from_notebook_node(self.notebook)
            self.result = output.encode("utf-8")
        elif self.operation == "Pdf":
            notebook = nbformat.reads(self.source.decode("utf-8-sig"), as_version=4)
            self.result, self.resources = self.exporter.from_notebook_node(notebook)
        else:
            raise ValueError(f"Unknown operation: {self.operation}")

    def finish(self, _request):
        if self.operation == "Parse":
            data = encode_json(self.result)
        elif self.operation == "SourceProjection":
            data = encode_json(self.result)
        elif isinstance(self.result, str):
            data = self.result.encode("utf-8")
        else:
            data = self.result
        self.artifact.write_bytes(data)
        asset_count = 0
        for name, content in (self.resources or {}).get("outputs", {}).items():
            if Path(name).name != name:
                raise ValueError(f"Unexpected exporter asset name: {name}")
            (self.artifact.parent / name).write_bytes(content)
            asset_count += 1
        return {"outputBytes": len(data), "assetCount": asset_count}


def worker():
    reference = ReferenceWorker()
    actions = {"prepare": reference.prepare, "run": reference.run, "finish": reference.finish}
    for line in sys.stdin:
        try:
            request = json.loads(line)
            response = actions[request["action"]](request)
            print(json.dumps({"ok": True, **(response or {})}), flush=True)
        except Exception:
            print(json.dumps({"ok": False, "error": traceback.format_exc()}), flush=True)


def environment(probe_chromium):
    versions = {}
    for name in ("nbformat", "nbconvert", "fastjsonschema", "playwright", "jinja2", "mistune", "pygments"):
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            pass
    if probe_chromium and "playwright" in versions:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as playwright:
            chromium = Path(playwright.chromium.executable_path)
            # Headless launches use the shell by default. Verify it can actually
            # start outside the timed samples rather than inferring from this path.
            try:
                with playwright.chromium.launch(headless=True, timeout=15000) as browser:
                    versions["chromium"] = browser.version
                    versions["chromiumPath"] = str(chromium)
            except Exception as error:
                versions["chromiumError"] = str(error).splitlines()[0]
    validator = os.environ.get("NBFORMAT_VALIDATOR", "fastjsonschema")
    print(json.dumps({"python": sys.version.split()[0], "validator": validator, **versions}))


def main():
    parser = argparse.ArgumentParser()
    modes = parser.add_subparsers(dest="mode", required=True)
    modes.add_parser("worker")
    modes.add_parser("environment").add_argument("--probe-chromium", action="store_true")
    generator = modes.add_parser("generate")
    generator.add_argument("path", type=Path)
    generator.add_argument("cells", type=int)
    generator.add_argument("code_chars", type=int)
    generator.add_argument("output_chars", type=int)
    generator.add_argument("image_pixels", type=int)
    arguments = parser.parse_args()
    if arguments.mode == "generate":
        generate(arguments.path, arguments.cells, arguments.code_chars, arguments.output_chars, arguments.image_pixels)
        print(hashlib.sha256(arguments.path.read_bytes()).hexdigest())
    elif arguments.mode == "environment":
        environment(arguments.probe_chromium)
    else:
        worker()


if __name__ == "__main__":
    main()
