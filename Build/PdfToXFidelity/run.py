"""PDF-to-X fidelity evaluation driver.

    python -I run.py corpus    --work <dir> --repo <checkout> [--tagged <pdf.js test/pdfs>]
    python -I run.py officeimo --work <dir> --tool <OfficeIMO.PdfToXFidelity.Tool.dll>
    python -I run.py tools     --work <dir> --venv-python <python.exe> [--only pdf2docx,docling,pymupdf4llm,libreoffice]
    python -I run.py normalize --work <dir>
    python -I run.py score     --work <dir>
    python -I run.py visual    --work <dir>
    python -I run.py report    --work <dir>

LibreOffice and Poppler run inside WSL (`wsl -e soffice`, `pdftoppm`, `pdfinfo`). Comparison tools run as child
processes from the `comparison` folder with paths passed as arguments.
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import csv
import glob
import json
import math
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from fidelity import model  # noqa: E402
from fidelity.markup import parse_struct_text, read_html, read_markdown  # noqa: E402
from fidelity.metrics import score  # noqa: E402
from fidelity.ooxml import read_docx, read_pptx, read_xlsx  # noqa: E402

WORD_ASSETS = ["BasicDocument.docx", "BasicDocumentWithSections.docx", "DocumentWithFootNotes.docx",
               "DocumentWithImages.docx", "DocumentWithTables.docx", "FieldsAndSections.docx",
               "Issue2510-SymbolBullets.docx", "List Library Bullet Styles.docx", "List Library Numbering Styles.docx",
               "ListsCreatedByWord.docx", "NestedTables.docx", "TableExamples.docx", "TestDocument365.docx",
               "sample1.docx", "sample2.docx", "sample3.docx", "ParagraphStyle.docx", "Normal Subscript Superscript.docx"]
APPLE_ASSETS = ["pages-created-v15.4.docx", "pages-created-tables-v15.4.docx", "pages-created-images-v15.4.docx",
                "pages-created-layout-v15.4.docx", "pages-created-numbering-v15.4.docx", "pages-everyday-v15.4.docx",
                "pages-list-continuation-v15.4.docx", "pages-paragraph-list-v15.4.docx", "pages-nine-level-saved-v15.4.docx",
                "pages-simple-v15.4.docx", "pages-table-borders-v15.4.docx", "pages-mixed-size-orders-v15.4.docx",
                "list-counter-explicit-v15.4.docx",
                "numbers-simple-v14.5.xlsx", "numbers-created-v14.5.xlsx", "numbers-formula-formatting-v14.5.xlsx",
                "numbers-table-borders-v14.5.xlsx", "numbers-banding-v14.5.xlsx", "numbers-cross-table-v14.5.xlsx",
                "excel-xlookup-optional-v16.113.3.xlsx",
                "keynote-created-rich-v15.4.pptx", "keynote-created-tables-v15.4.pptx", "keynote-created-images-v15.4.pptx",
                "keynote-simple-v15.4.pptx", "keynote-slide-links-v15.4.pptx", "keynote-table-borders-v15.4.pptx",
                "keynote-colors-v15.4.pptx", "list-counter-explicit-v15.4.pptx"]
EXCEL_ASSETS = ["BasicExcel.xlsx", "SampleFileExcel.xlsx"]

OFFICEIMO_TARGETS = ["docx", "odt", "rtf", "html", "md", "xlsx", "ods", "pptx", "odp"]
# Which targets are scored for each source family. Semantic targets apply to everything.
FAMILY_TARGETS = {"word": ["docx", "odt", "rtf", "html", "md", "xlsx", "ods"],
                  "tagged": ["docx", "odt", "rtf", "html", "md"],
                  "excel": ["xlsx", "ods", "html", "md", "docx"],
                  "slides": ["pptx", "odp", "html", "md"]}
NORMALIZE = {"odt": "docx", "rtf": "docx", "ods": "xlsx", "odp": "pptx"}
VISUAL_TARGETS = ["docx", "odt", "rtf", "xlsx", "ods", "pptx", "odp"]


def wsl_path(path: str) -> str:
    path = os.path.abspath(path)
    drive, rest = os.path.splitdrive(path)
    return "/mnt/" + drive[0].lower() + rest.replace("\\", "/")


def soffice(files: list[str], outdir: str, convert_to: str, infilter: str | None = None, timeout: int = 1800) -> None:
    """Batch conversion through LibreOffice in WSL with a private profile."""
    if not files:
        return
    os.makedirs(outdir, exist_ok=True)
    for start in range(0, len(files), 40):
        cmd = ["wsl", "-e", "soffice", "--headless", "--norestore", "-env:UserInstallation=file:///tmp/pdfx-lo-profile"]
        if infilter:
            cmd.append(f"--infilter={infilter}")
        cmd += ["--convert-to", convert_to, "--outdir", wsl_path(outdir)] + [wsl_path(f) for f in files[start:start + 40]]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=timeout, check=False)


def load_manifest(work: str) -> list[dict]:
    with open(os.path.join(work, "corpus", "manifest.json"), encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------------------------------------------------ corpus

def cmd_corpus(args) -> None:
    work = os.path.abspath(args.work)
    corpus = os.path.join(work, "corpus")
    src_dir = os.path.join(corpus, "sources")
    pdf_dir = os.path.join(corpus, "pdf")
    os.makedirs(src_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    docs = os.path.join(args.repo, "OfficeIMO.TestAssets", "Documents")
    entries: list[dict] = []

    def add(source: str, set_name: str, licence: str) -> None:
        stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", os.path.splitext(os.path.basename(source))[0]).strip("-")
        ext = os.path.splitext(source)[1].lower()
        family = {".docx": "word", ".xlsx": "excel", ".pptx": "slides"}[ext]
        doc_id = f"{set_name}--{stem}{'-' + ext[1:] if ext != '.docx' else ''}"
        dest = os.path.join(src_dir, doc_id + ext)
        shutil.copyfile(source, dest)
        entries.append({"id": doc_id, "family": family, "source": dest, "pdf": os.path.join(pdf_dir, doc_id + ".pdf"),
                        "producer": "libreoffice", "set": set_name, "licence": licence})

    generated = os.path.join(work, "tmp", "generated")
    subprocess.run([args.venv_python, "-I", os.path.join(HERE, "generate_sources.py"), generated], check=True)
    for path in sorted(glob.glob(os.path.join(generated, "*.*"))):
        add(path, "generated", "generated by this harness")
    for name in WORD_ASSETS + EXCEL_ASSETS:
        add(os.path.join(docs, name), "testassets", "OfficeIMO repository test assets (MIT)")
    for name in APPLE_ASSETS:
        add(os.path.join(docs, "IWorkCorpus", "native-exports", name), "iwork-exports",
            "OfficeIMO IWorkCorpus native exports (see IWorkCorpus/THIRD-PARTY-NOTICES.md)")
    for path in sorted(glob.glob(os.path.join(docs, "WordPageFootnotes", "*.docx"))):
        add(path, "word-footnotes", "OfficeIMO WordPageFootnotes fixtures (MIT)")
    # LibreOffice is the independent PDF producer for every source.
    soffice([e["source"] for e in entries if not os.path.exists(e["pdf"])], pdf_dir, "pdf")
    # The footnote fixtures also carry Microsoft Word PDFs: a second producer for the same sources.
    for path in sorted(glob.glob(os.path.join(docs, "WordPageFootnotes", "*.microsoft-word.pdf"))):
        stem = os.path.basename(path).replace(".microsoft-word.pdf", "")
        base = next((e for e in entries if e["set"] == "word-footnotes" and e["id"].endswith("--" + stem)), None)
        if base:
            doc_id = base["id"].replace("word-footnotes--", "word-footnotes-msword--")
            dest = os.path.join(pdf_dir, doc_id + ".pdf")
            shutil.copyfile(path, dest)
            entries.append(dict(base, id=doc_id, pdf=dest, producer="microsoft-word", set="word-footnotes-msword"))
    if args.tagged:
        struct_dir = os.path.join(corpus, "struct")
        os.makedirs(struct_dir, exist_ok=True)
        for path in sorted(glob.glob(os.path.join(args.tagged, "*.pdf"))):
            with open(path, "rb") as f:
                head = f.read()
            if b"StructTreeRoot" not in head or len(head) > 4_000_000:
                continue
            try:
                proc = subprocess.run(["wsl", "-e", "pdfinfo", "-struct-text", wsl_path(path)], capture_output=True, timeout=30)
            except subprocess.TimeoutExpired:
                continue
            text = proc.stdout.decode("utf-8", "replace")
            gold = parse_struct_text(text)
            if len(gold.reading_words()) < 30 or not gold.blocks:
                continue
            doc_id = "pdfjs-tagged--" + os.path.splitext(os.path.basename(path))[0]
            dest = os.path.join(pdf_dir, doc_id + ".pdf")
            shutil.copyfile(path, dest)
            with open(os.path.join(struct_dir, doc_id + ".txt"), "w", encoding="utf-8") as f:
                f.write(text)
            entries.append({"id": doc_id, "family": "tagged", "source": os.path.join(struct_dir, doc_id + ".txt"),
                            "pdf": dest, "producer": "original", "set": "pdfjs-tagged",
                            "licence": "pdf.js test corpus (Apache-2.0 repository; individual files keep their origin)"})
    entries = [e for e in entries if os.path.exists(e["pdf"])]
    with open(os.path.join(corpus, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=1)
    print(f"{len(entries)} documents")


# ------------------------------------------------------------------------------------------------ conversions

def cmd_officeimo(args) -> None:
    work = os.path.abspath(args.work)
    out = os.path.join(work, "outputs", "officeimo")
    os.makedirs(out, exist_ok=True)
    entries = load_manifest(work)

    def one(e):
        targets = ",".join(t for t in OFFICEIMO_TARGETS if t in FAMILY_TARGETS[e["family"]])
        try:
            subprocess.run(["dotnet", args.tool, "convert", e["pdf"], out, targets], timeout=args.timeout,
                           stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, check=False)
            return e["id"], "ok"
        except subprocess.TimeoutExpired:
            return e["id"], "timeout"

    with cf.ThreadPoolExecutor(args.parallel) as pool:
        for doc_id, status in pool.map(one, entries):
            if status != "ok":
                with open(os.path.join(out, doc_id + ".officeimo.json"), "w", encoding="utf-8") as f:
                    json.dump([{"target": t, "status": "timeout"} for t in OFFICEIMO_TARGETS], f)
            print(doc_id, status, flush=True)


def cmd_tools(args) -> None:
    work = os.path.abspath(args.work)
    entries = load_manifest(work)
    only = set(args.only.split(",")) if args.only else {"pdf2docx", "docling", "pymupdf4llm", "libreoffice"}
    scripts = os.path.join(HERE, "comparison")
    env = dict(os.environ, HF_HOME=os.path.join(work, "tmp", "hf"), TEMP=os.path.join(work, "tmp"), TMP=os.path.join(work, "tmp"))
    jobs = os.path.join(work, "tmp", "jobs")
    os.makedirs(jobs, exist_ok=True)

    def batch(tool: str, families: set[str]) -> str:
        rows = [{"id": e["id"], "pdf": e["pdf"]} for e in entries if e["family"] in families]
        path = os.path.join(jobs, tool + ".json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(rows, f)
        return path

    word_like = {"word", "tagged", "excel"}
    if "pdf2docx" in only:
        out = os.path.join(work, "outputs", "pdf2docx")
        subprocess.run([args.venv_python, "-I", os.path.join(scripts, "run_pdf2docx.py"), batch("pdf2docx", word_like), out],
                       env=env, check=False)
    if "pymupdf4llm" in only:
        out = os.path.join(work, "outputs", "pymupdf4llm")
        subprocess.run([args.venv_python, "-I", os.path.join(scripts, "run_pymupdf4llm.py"),
                        batch("pymupdf4llm", set(FAMILY_TARGETS)), out], env=env, check=False)
    if "docling" in only:
        out = os.path.join(work, "outputs", "docling")
        subprocess.run([args.venv_python, "-I", os.path.join(scripts, "run_docling.py"), batch("docling", set(FAMILY_TARGETS)), out],
                       env=env, check=False)
    if "libreoffice" in only:
        out = os.path.join(work, "outputs", "libreoffice")
        os.makedirs(out, exist_ok=True)
        soffice([e["pdf"] for e in entries if e["family"] in word_like], out, "docx:MS Word 2007 XML", "writer_pdf_import")
        soffice([e["pdf"] for e in entries if e["family"] == "slides"], out, "pptx", "impress_pdf_import")


def cmd_normalize(args) -> None:
    """Opens OpenDocument and RTF outputs in LibreOffice and saves them as OOXML for one shared reader."""
    work = os.path.abspath(args.work)
    src = os.path.join(work, "outputs", "officeimo")
    norm = os.path.join(src, "normalized")
    for ext, target in NORMALIZE.items():
        files = sorted(glob.glob(os.path.join(src, "*." + ext)))
        soffice(files, os.path.join(norm, ext), target)


# ------------------------------------------------------------------------------------------------ scoring

def read_output(path: str, ext: str) -> model.Document:
    if ext in ("docx",):
        return read_docx(path)
    if ext == "xlsx":
        return read_xlsx(path)
    if ext == "pptx":
        return read_pptx(path)
    if ext == "html":
        return read_html(path)
    if ext == "md":
        return read_markdown(path)
    raise ValueError(ext)


def read_gold(e: dict) -> model.Document:
    if e["family"] == "tagged":
        with open(e["source"], encoding="utf-8") as f:
            return parse_struct_text(f.read())
    return {"word": read_docx, "excel": read_xlsx, "slides": read_pptx}[e["family"]](e["source"])


def outputs_for(work: str, e: dict) -> list[tuple[str, str, str | None, str]]:
    """(tool, target, path or None, status) for every scored conversion of one document."""
    rows = []
    oi = os.path.join(work, "outputs", "officeimo")
    status = {}
    meta = os.path.join(oi, e["id"] + ".officeimo.json")
    if os.path.exists(meta):
        with open(meta, encoding="utf-8") as f:
            status = {r["target"]: r for r in json.load(f)}
    for target in FAMILY_TARGETS[e["family"]]:
        s = status.get(target, {}).get("status", "missing")
        if target in NORMALIZE:
            path = os.path.join(oi, "normalized", target, e["id"] + "." + NORMALIZE[target])
            if s == "ok" and not os.path.exists(path):
                s = "unreadable"
        else:
            path = os.path.join(oi, e["id"] + "." + target)
        rows.append(("officeimo", target, path if s == "ok" and os.path.exists(path) else None, s))
    comparisons = [("pdf2docx", "docx", {"word", "tagged", "excel"}), ("libreoffice", "docx", {"word", "tagged", "excel"}),
                   ("libreoffice", "pptx", {"slides"}), ("docling", "md", set(FAMILY_TARGETS)),
                   ("docling", "html", set(FAMILY_TARGETS)), ("pymupdf4llm", "md", set(FAMILY_TARGETS))]
    for tool, target, families in comparisons:
        if e["family"] not in families or target not in FAMILY_TARGETS[e["family"]]:
            continue
        folder = os.path.join(work, "outputs", tool)
        if not os.path.isdir(folder):
            continue
        path = os.path.join(folder, e["id"] + "." + target)
        rows.append((tool, target, path if os.path.exists(path) else None, "ok" if os.path.exists(path) else "error"))
    return rows


def _score_one(work: str, e: dict) -> list[dict]:
    gold = read_gold(e)
    results = []
    for tool, target, path, status in outputs_for(work, e):
        ext = NORMALIZE.get(target, target)
        out = model.Document()
        if path:
            try:
                out = read_output(path, ext)
            except Exception as exc:  # unreadable output counts as an empty conversion
                status = "unreadable: " + type(exc).__name__
        row = {"id": e["id"], "set": e["set"], "family": e["family"], "producer": e["producer"], "tool": tool,
               "target": target, "status": status}
        row.update(score(gold, out))
        results.append(row)
    return results


def cmd_score(args) -> None:
    work = os.path.abspath(args.work)
    entries = load_manifest(work)
    rows = []
    with cf.ProcessPoolExecutor(args.parallel) as pool:
        futures = {pool.submit(_score_one, work, e): e["id"] for e in entries}
        for future in cf.as_completed(futures):
            try:
                rows.extend(future.result())
            except Exception as exc:
                print("score failed", futures[future], exc, flush=True)
    rows.sort(key=lambda r: (r["id"], r["tool"], r["target"]))
    write_csv(os.path.join(work, "results.csv"), rows)
    print(f"{len(rows)} scored conversions")


def write_csv(path: str, rows: list[dict]) -> None:
    keys: list[str] = []
    for r in rows:
        for k in r:
            if k not in keys:
                keys.append(k)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, keys)
        w.writeheader()
        for r in rows:
            w.writerow({k: (round(v, 4) if isinstance(v, float) and not math.isnan(v) else ("" if isinstance(v, float) else v))
                        for k, v in r.items()})


# ------------------------------------------------------------------------------------------------ visual

def cmd_visual(args) -> None:
    from fidelity.visual import compare_pdfs

    work = os.path.abspath(args.work)
    entries = {e["id"]: e for e in load_manifest(work)}
    vis = os.path.join(work, "tmp", "visual")
    jobs = []
    for tool in ("officeimo", "pdf2docx", "libreoffice"):
        folder = os.path.join(work, "outputs", tool)
        for target in VISUAL_TARGETS:
            files = [p for p in glob.glob(os.path.join(folder, "*." + target)) if os.path.basename(p).rsplit(".", 1)[0] in entries]
            if not files:
                continue
            out = os.path.join(vis, tool, target)
            soffice([p for p in files if not os.path.exists(os.path.join(out, os.path.basename(p).rsplit(".", 1)[0] + ".pdf"))],
                    out, "pdf")
            for p in files:
                doc_id = os.path.basename(p).rsplit(".", 1)[0]
                jobs.append((tool, target, doc_id, entries[doc_id]["pdf"], os.path.join(out, doc_id + ".pdf")))
    rows = []
    with cf.ThreadPoolExecutor(args.parallel) as pool:
        for job, res in zip(jobs, pool.map(lambda j: compare_pdfs(j[3], j[4], os.path.join(vis, "png")), jobs)):
            rows.append({"id": job[2], "tool": job[0], "target": job[1], **res})
    write_csv(os.path.join(work, "visual.csv"), rows)
    print(f"{len(rows)} visual comparisons")


def cmd_report(args) -> None:
    from fidelity.report import write_report

    write_report(os.path.abspath(args.work))


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("command", choices=["corpus", "officeimo", "tools", "normalize", "score", "visual", "report"])
    p.add_argument("--work", required=True)
    p.add_argument("--repo", default=os.path.abspath(os.path.join(HERE, "..", "..")))
    p.add_argument("--tagged")
    p.add_argument("--tool")
    p.add_argument("--venv-python", default=sys.executable)
    p.add_argument("--only")
    p.add_argument("--parallel", type=int, default=8)
    p.add_argument("--timeout", type=int, default=300)
    args = p.parse_args()
    {"corpus": cmd_corpus, "officeimo": cmd_officeimo, "tools": cmd_tools, "normalize": cmd_normalize,
     "score": cmd_score, "visual": cmd_visual, "report": cmd_report}[args.command](args)


if __name__ == "__main__":
    main()
