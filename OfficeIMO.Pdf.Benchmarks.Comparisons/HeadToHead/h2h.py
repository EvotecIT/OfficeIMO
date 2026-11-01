#!/usr/bin/env python3
"""Corpus, deploy-size, fidelity and contact-sheet tooling for the HTML head-to-head lane.

Needs Python 3 with Pillow, plus pdftoppm, pdftotext, pdfinfo, pdffonts and qpdf on PATH.
"""
import argparse
import base64
import collections
import difflib
import json
import os
import random
import re
import shutil
import statistics
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

HERE = Path(__file__).resolve().parent
LANE = HERE.parent
ENGINES = ["Chromium", "OfficeIMO", "PeachPDF", "PdfPecker"]
DPI = 96

# Repository fixtures: the OfficeIMO qualification set, so reported separately from authored pages.
REPO_FIXTURES = {
    "repo-invoice": "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/representative/invoice.html",
    "repo-quarterly-report": "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/representative/quarterly-report.html",
    "repo-dashboard-print": "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/representative/dashboard-print.html",
    "repo-multilingual-bidi": "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/representative/multilingual-bidi.html",
    "repo-subgrid-operations": "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/advanced-held-out/subgrid-operations.html",
}
# Peer samples under MIT (PdfPecker) and BSD-style (PeachPDF) terms. They are read from a local
# clone at run time and never committed, so their notices stay with their own repositories.
PEER_FIXTURES = {
    "peer-pdfpecker-invoice": "PdfPecker/tests/PdfPecker.Tests/Golden/complex/invoice.html",
    "peer-pdfpecker-flex-grid": "PdfPecker/tests/PdfPecker.Tests/Golden/complex/flex-grid.html",
    "peer-pdfpecker-report": "PdfPecker/tests/PdfPecker.Tests/Golden/complex/report.html",
    "peer-peachpdf-tables": "PeachPDF/src/PeachPDF.TestHarness/tables.htm",
}
COLR_FONT = "PeachPDF/assets/fonts/NotoColorEmoji-Subset.ttf"


def report_html():
    rng = random.Random(20261101)
    regions = ["North", "South", "East", "West", "Central", "Coastal"]
    products = ["Bearings", "Valves", "Filters", "Gaskets", "Pumps", "Seals", "Hoses", "Couplings"]
    words = ("the review found steady demand across most regions while supply remained tight during the "
             "second half of the year and several contracts were renewed on improved terms after careful "
             "negotiation with long standing customers who value reliable delivery").split()
    out = ['<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Operations report</title><style>',
           '@page{size:A4;margin:18mm 16mm;@bottom-right{content:"Page " counter(page) " of " counter(pages);font-size:8pt;color:#567}}',
           'body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;line-height:1.35;color:#1c2230}',
           'h1{font-size:24pt;margin:0 0 6pt}h2{font-size:15pt;color:#123a6b;break-before:page;margin:0 0 6pt}',
           'table{width:100%;border-collapse:collapse;font-size:8.5pt;margin:8pt 0}th{background:#123a6b;color:#fff;text-align:left;padding:3pt 4pt}',
           'td{border-bottom:1px solid #d5dbe3;padding:2.5pt 4pt}td.n,th.n{text-align:right}tbody tr:nth-child(even){background:#f3f6f8}',
           'tfoot td{font-weight:bold;border-top:1.5pt solid #123a6b}.kpi{display:flex;gap:8pt;margin:8pt 0}.kpi div{flex:1;background:#eef3f8;padding:6pt;border-radius:4pt}',
           '.kpi b{display:block;font-size:14pt;color:#0b6e6e}tr{break-inside:avoid}</style></head><body>',
           '<h1>Operations report 2026</h1><p>Prepared for the executive committee. All figures are illustrative and generated deterministically.</p>']
    out.append('<table><thead><tr><th>Chapter</th><th>Topic</th></tr></thead><tbody>')
    chapters = 15
    for chapter in range(1, chapters + 1):
        out.append(f'<tr><td>{chapter}</td><td>{regions[chapter % 6]} operations review</td></tr>')
    out.append('</tbody></table>')
    for chapter in range(1, chapters + 1):
        out.append(f'<h2 id="c{chapter}">{chapter}. {regions[chapter % 6]} operations review</h2>')
        out.append('<div class="kpi">' + ''.join(
            f'<div><b>{rng.randint(70, 130)}%</b>{label}</div>' for label in ("of plan", "on time", "utilised")) + '</div>')
        for _ in range(2):
            out.append('<p>' + ' '.join(rng.choice(words) for _ in range(70)).capitalize() + '.</p>')
        rows = 56 if chapter < chapters else 30
        out.append('<table><thead><tr><th>Order</th><th>Product</th><th>Region</th><th class="n">Units</th>'
                   '<th class="n">Unit price</th><th class="n">Net</th><th class="n">Margin</th></tr></thead><tbody>')
        total = 0.0
        for row in range(rows):
            units = rng.randint(5, 900)
            price = rng.randint(120, 9800) / 100
            net = units * price
            total += net
            out.append(f'<tr><td>SO-{chapter:02d}{row:04d}</td><td>{rng.choice(products)}</td><td>{rng.choice(regions)}</td>'
                       f'<td class="n">{units:,}</td><td class="n">{price:,.2f}</td><td class="n">{net:,.2f}</td>'
                       f'<td class="n">{rng.randint(4, 38)}%</td></tr>')
        out.append(f'</tbody><tfoot><tr><td colspan="5">Chapter total</td><td class="n">{total:,.2f}</td><td></td></tr></tfoot></table>')
    out.append('</body></html>')
    return '\n'.join(out)


def command_corpus(args):
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    provenance = {}
    for page in sorted((LANE / "Corpus" / "HeadToHead").glob("*.html")):
        shutil.copy(page, out / page.name)
        provenance[page.stem] = "authored"
    (out / "report-30p.html").write_text(report_html(), encoding="utf-8")
    provenance["report-30p"] = "authored"
    repo = LANE.parent
    for name, relative in REPO_FIXTURES.items():
        shutil.copy(repo / relative, out / (name + ".html"))
        provenance[name] = "repository"
    if args.peers:
        peers = Path(args.peers)
        for name, relative in PEER_FIXTURES.items():
            shutil.copy(peers / relative, out / (name + ".html"))
            provenance[name] = "peer"
        font = peers / COLR_FONT
        encoded = base64.b64encode(font.read_bytes()).decode("ascii")
        (out / "peer-colr-font.html").write_text(
            '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>COLR font</title><style>'
            f'@font-face{{font-family:Colr;src:url(data:font/ttf;base64,{encoded})}}'
            '@page{size:A4;margin:18mm}body{font-family:Arial,sans-serif}p{font-family:Colr,sans-serif;font-size:40pt}</style>'
            '</head><body><h1>Web font with colour glyphs</h1><p>😀 🎉 🚀 ❤️ 👍</p></body></html>', encoding="utf-8")
        provenance["peer-colr-font"] = "peer"
    (out / "provenance.json").write_text(json.dumps(provenance, indent=1, sort_keys=True))
    print(f"{len(provenance)} documents in {out}")


DEPLOY = {
    "OfficeIMO": ('<ProjectReference Include="{repo}/OfficeIMO.Html/OfficeIMO.Html.csproj" />',
                  'using OfficeIMO.Html;using OfficeIMO.Html.Pdf;'
                  'using var d=HtmlConversionDocument.Parse("<p>x</p>").ToPdf();Console.WriteLine(d.ToBytes().Length);'),
    "PeachPDF": ('<PackageReference Include="PeachPDF" Version="{peach}" />',
                 'var d=await new PeachPDF.PdfGenerator().GeneratePdf("<p>x</p>",new PeachPDF.PdfGenerateConfig{PageSize=PeachPDF.PageSize.A4});'
                 'using var s=new MemoryStream();d.Save(s);Console.WriteLine(s.Length);'),
    "PdfPecker": ('<PackageReference Include="PdfPecker" Version="{pecker}" />',
                  'Console.WriteLine(PdfPecker.HtmlToPdf.Convert("<p>x</p>").Length);'),
    "Baseline": ('', 'Console.WriteLine(0);'),
}


def tree_size(path):
    return sum(item.stat().st_size for item in Path(path).rglob("*") if item.is_file())


def command_deploy(args):
    project = (LANE / "OfficeIMO.Pdf.Benchmarks.Comparisons.csproj").read_text()
    peach = re.search(r"<PeachPdfBenchmarkVersion>([^<]+)", project).group(1)
    pecker = re.search(r"<PdfPeckerBenchmarkVersion>([^<]+)", project).group(1)
    root = Path(args.out)
    results = {}
    for engine, (reference, program) in DEPLOY.items():
        directory = root / engine
        shutil.rmtree(directory, ignore_errors=True)
        directory.mkdir(parents=True)
        reference = reference.format(repo=LANE.parent, peach=peach, pecker=pecker)
        (directory / "Probe.csproj").write_text(
            '<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net10.0</TargetFramework>'
            '<ImplicitUsings>enable</ImplicitUsings><Nullable>enable</Nullable><SatelliteResourceLanguages>en</SatelliteResourceLanguages>'
            f'</PropertyGroup><ItemGroup>{reference}</ItemGroup></Project>')
        (directory / "Program.cs").write_text(program)
        shutil.copy(LANE.parent / "global.json", directory / "global.json")
        publish = directory / "out"
        subprocess.run(["dotnet", "publish", "-c", "Release", "-r", args.rid, "--self-contained", "false", "-o", str(publish), "-v", "q", "--nologo"],
                       cwd=directory, check=True, stdout=subprocess.DEVNULL)
        files = sorted(((f.stat().st_size, f.name) for f in publish.rglob("*") if f.is_file() and f.suffix not in (".pdb", ".xml")), reverse=True)
        results[engine] = {"bytes": sum(size for size, _ in files), "files": len(files), "largest": files[:6]}
        print(engine, results[engine]["bytes"])
    cache = Path.home() / "Library/Caches/ms-playwright"
    shells = sorted(cache.glob("chromium_headless_shell-*"))
    if shells:
        results["Chromium"] = {"bytes": tree_size(shells[-1]), "files": 0,
                               "largest": [[0, "chromium headless shell only; excludes HtmlTinkerX, Playwright and its node driver"]]}
    (root / "deploy-sizes.json").write_text(json.dumps(results, indent=1))


def run(command):
    return subprocess.run(command, capture_output=True, text=True, errors="replace")


def rasterise(pdf, directory):
    directory.mkdir(parents=True, exist_ok=True)
    for old in directory.glob("*.png"):
        old.unlink()
    run(["pdftoppm", "-r", str(DPI), "-png", str(pdf), str(directory / "p")])
    return sorted(directory.glob("*.png"), key=lambda p: int(re.search(r"(\d+)\.png$", p.name).group(1)))


def mismatch_pixels(a, b, threshold):
    channels = ImageChops.difference(a, b).split()
    worst = ImageChops.lighter(ImageChops.lighter(channels[0], channels[1]), channels[2])
    return sum(worst.point(lambda value: 255 if value >= threshold else 0).histogram()[255:]), worst


def ink_pixels(image):
    gray = image.convert("L").point(lambda value: 255 if value < 245 else 0)
    return gray


def compare_page(left, right):
    """Exact mismatch, plus an ink-normalised score on blurred pages so antialiasing and
    sub-pixel glyph differences do not dominate; layout displacement still counts."""
    a = Image.open(left).convert("RGB")
    b = Image.open(right).convert("RGB")
    if a.size != b.size:
        canvas = Image.new("RGB", a.size, "white")
        canvas.paste(b, (0, 0))
        b = canvas
    exact, _ = mismatch_pixels(a, b, 32)
    blurred_a = a.filter(ImageFilter.GaussianBlur(2.5))
    blurred_b = b.filter(ImageFilter.GaussianBlur(2.5))
    blurred, _ = mismatch_pixels(blurred_a, blurred_b, 40)
    union = ImageChops.lighter(ink_pixels(blurred_a), ink_pixels(blurred_b)).histogram()[255]
    total = a.size[0] * a.size[1]
    return {"exact": exact / total, "ink": min(1.0, blurred / max(union, 1)), "union": union / total}


def page_ink(path):
    image = Image.open(path).convert("RGB")
    return ink_pixels(image).histogram()[255] / (image.size[0] * image.size[1])


def tokens(pdf):
    text = run(["pdftotext", "-enc", "UTF-8", str(pdf), "-"]).stdout
    found = re.findall(r"[0-9A-Za-zÀ-ɏ]+|[^\s0-9A-Za-zÀ-ɏ]", text)
    return found


def compare_text(reference, candidate):
    common = sum((collections.Counter(reference) & collections.Counter(candidate)).values())
    recall = common / max(len(reference), 1)
    limit = 6000  # SequenceMatcher is quadratic in the worst case; the head covers a 30-page report's structure.
    order = difflib.SequenceMatcher(None, reference[:limit], candidate[:limit], autojunk=False).ratio() if reference else 0
    return recall, order


def pdf_facts(pdf):
    info = run(["pdfinfo", str(pdf)]).stdout
    pages = int((re.search(r"Pages:\s+(\d+)", info) or [0, 0])[1])
    size = (re.search(r"Page size:\s+(.+)", info) or ["", ""])[1].strip()
    tagged = (re.search(r"Tagged:\s+(\w+)", info) or ["", "?"])[1]
    font_lines = run(["pdffonts", str(pdf)]).stdout.splitlines()[2:]
    fonts = sorted({line.split()[0].split("+")[-1] for line in font_lines if line.strip()})
    check = run(["qpdf", "--check", str(pdf)])
    return {"pages": pages, "size": size, "tagged": tagged, "fonts": fonts, "qpdf": check.returncode, "qpdfNote": check.stderr.strip()[:200]}


def medians(records):
    result = {}
    by_key = collections.defaultdict(list)
    for record in records:
        by_key[(record["document"], record["engine"])].append(record)
    for key, items in by_key.items():
        cold = [r for r in items if r["mode"] == "cold" and r["status"] == "Passed"]
        warm = [r for r in items if r["mode"] == "warm" and r["status"] == "Passed"]
        parallel = [r for r in items if r["mode"] == "parallel" and r["status"] == "Passed"]
        failed = [r for r in items if r["status"] != "Passed"]
        result[key] = {
            "status": failed[0]["status"] + ": " + failed[0]["detail"] if failed else "Passed",
            "firstMs": statistics.median(r["firstMilliseconds"] for r in cold) if cold else None,
            "firstRuns": len(cold),
            "firstCpuMs": statistics.median(r["firstCpuMilliseconds"] for r in cold) if cold else None,
            "warmCpuMs": statistics.median(r["warmCpuMedianMilliseconds"] for r in warm) if warm else None,
            "warmMinMs": min(r["warmMinimumMilliseconds"] for r in warm) if warm else None,
            "warmMs": statistics.median(r["warmMedianMilliseconds"] for r in warm) if warm else None,
            "allocMb": statistics.median(r["warmAllocatedBytes"] for r in warm) / 1e6 if warm else None,
            "firstAllocMb": statistics.median(r["firstAllocatedBytes"] for r in cold) / 1e6 if cold else None,
            "peakMb": statistics.median(r["peakRssBytes"] for r in cold) / 1e6 if cold else None,
            "warmPeakMb": statistics.median(r["peakRssBytes"] for r in warm) / 1e6 if warm else None,
            "docsPerSecond": statistics.median(r["parallelDocumentsPerSecond"] for r in parallel) if parallel else None,
            "outputKb": cold[-1]["outputBytes"] / 1024 if cold else None,
        }
    return result


def command_analyze(args):
    run_dir = Path(args.run)
    corpus = Path(args.corpus)
    records = [json.loads(line) for line in (run_dir / "measurements.jsonl").read_text().splitlines() if line.strip()]
    timing = medians(records)
    provenance = json.loads((corpus / "provenance.json").read_text())
    documents = sorted(provenance)
    raster_root = run_dir / "raster"
    rows = []
    for document in documents:
        reference_pdf = run_dir / "Chromium" / (document + ".pdf")
        reference_pages = rasterise(reference_pdf, raster_root / "Chromium" / document) if reference_pdf.exists() else []
        reference_tokens = tokens(reference_pdf) if reference_pdf.exists() else []
        for engine in ENGINES:
            row = {"document": document, "provenance": provenance[document], "engine": engine}
            row.update(timing.get((document, engine), {"status": "Not run"}))
            pdf = run_dir / engine / (document + ".pdf")
            if pdf.exists() and row["status"] == "Passed":
                row.update(pdf_facts(pdf))
                pages = reference_pages if engine == "Chromium" else rasterise(pdf, raster_root / engine / document)
                if reference_pages and engine != "Chromium":
                    scores = [compare_page(x, y) for x, y in zip(reference_pages, pages)]
                    extra = reference_pages[len(pages):] or pages[len(reference_pages):]
                    # Missing or surplus pages count in full where they hold ink.
                    inks = [s["ink"] for s in scores] + [1.0 for page in extra if page_ink(page) > 0.001]
                    row["inkDiff"] = statistics.mean(inks) if inks else None
                    row["exactDiff"] = statistics.mean(s["exact"] for s in scores) if scores else None
                    row["worstPage"] = max(range(len(scores)), key=lambda i: scores[i]["ink"]) + 1 if scores else None
                    row["pageDelta"] = len(pages) - len(reference_pages)
                    row["textRecall"], row["textOrder"] = compare_text(reference_tokens, tokens(pdf))
            rows.append(row)
            print(document, engine, row.get("pages"), row.get("inkDiff"), flush=True)
    (run_dir / "analysis.json").write_text(json.dumps(rows, indent=1))
    write_tables(run_dir, rows)


def fmt(value, pattern="{:.0f}", missing="—"):
    return missing if value is None else pattern.format(value)


def write_tables(run_dir, rows):
    lines = ["| Document | Engine | Pages | First ms | First CPU ms | Warm ms | Warm CPU ms | Alloc MB | Peak MB | Par docs/s | KB | Ink diff | Exact diff | Text recall | Text order |",
             "|---|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|"]
    for row in rows:
        if row["status"] != "Passed":
            lines.append(f"| {row['document']} | {row['engine']} | **{row['status'][:70]}** |" + " |" * 12)
            continue
        lines.append("| " + " | ".join([
            row["document"], row["engine"], fmt(row.get("pages")), fmt(row.get("firstMs")), fmt(row.get("firstCpuMs")), fmt(row.get("warmMs"), "{:.1f}"), fmt(row.get("warmCpuMs"), "{:.1f}"),
            fmt(row.get("allocMb"), "{:.1f}"), fmt(row.get("peakMb")), fmt(row.get("docsPerSecond"), "{:.1f}"), fmt(row.get("outputKb")),
            fmt(row.get("inkDiff"), "{:.1%}"), fmt(row.get("exactDiff"), "{:.2%}"), fmt(row.get("textRecall"), "{:.1%}"), fmt(row.get("textOrder"), "{:.1%}")]) + " |")
    (run_dir / "results.md").write_text("\n".join(lines) + "\n")


def command_sheet(args):
    run_dir = Path(args.run)
    raster = run_dir / "raster"
    sheets = run_dir / "contact"
    sheets.mkdir(exist_ok=True)
    rows = json.loads((run_dir / "analysis.json").read_text())
    documents = sorted({row["document"] for row in rows})
    by = {(row["document"], row["engine"]): row for row in rows}
    width, height = 470, 665
    strips = []
    for document in documents:
        reference = sorted((raster / "Chromium" / document).glob("*.png"))
        worst = max((by[(document, e)].get("worstPage") or 1 for e in ENGINES[1:]), default=1)
        chosen = [1] if len(reference) < 2 else sorted({1, worst})
        for page in chosen:
            strip = Image.new("RGB", (len(ENGINES) * (width + 8), height + 40), "#e8e8e8")
            draw = ImageDraw.Draw(strip)
            for column, engine in enumerate(ENGINES):
                pages = sorted((raster / engine / document).glob("*.png"), key=lambda p: int(re.search(r"(\d+)\.png$", p.name).group(1)))
                row = by[(document, engine)]
                ink = row.get("inkDiff")
                label = f"{document} p{page} | {engine} | {row.get('pages', 'FAILED')}pp" + (f" | ink {ink:.0%}" if ink is not None else "")
                draw.text((column * (width + 8) + 4, 6), label, fill="black")
                x = column * (width + 8)
                if page <= len(pages):
                    image = Image.open(pages[page - 1]).convert("RGB")
                    image.thumbnail((width, height))
                    strip.paste(image, (x, 28))
                else:
                    draw.text((x + 30, 200), "no page" if pages else "no output: " + str(row["status"])[:60], fill="#b00")
            path = sheets / f"{document}-p{page}.png"
            strip.save(path)
            strips.append(path)
    gallery = ['<!doctype html><meta charset="utf-8"><title>HTML head-to-head contact sheet</title><style>body{font-family:sans-serif;background:#ddd}img{width:100%;display:block;margin-bottom:10px}</style>']
    gallery += [f'<img loading="lazy" src="{path.name}" alt="{path.stem}">' for path in strips]
    (sheets / "index.html").write_text("\n".join(gallery))
    print(sheets / "index.html")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    corpus = commands.add_parser("corpus")
    corpus.add_argument("--out", required=True)
    corpus.add_argument("--peers")
    corpus.set_defaults(handler=command_corpus)
    deploy = commands.add_parser("deploy")
    deploy.add_argument("--out", required=True)
    deploy.add_argument("--rid", default="osx-arm64")
    deploy.set_defaults(handler=command_deploy)
    analyze = commands.add_parser("analyze")
    analyze.add_argument("--run", required=True)
    analyze.add_argument("--corpus", required=True)
    analyze.set_defaults(handler=command_analyze)
    sheet = commands.add_parser("sheet")
    sheet.add_argument("--run", required=True)
    sheet.set_defaults(handler=command_sheet)
    args = parser.parse_args()
    args.handler(args)


if __name__ == "__main__":
    sys.exit(main())
