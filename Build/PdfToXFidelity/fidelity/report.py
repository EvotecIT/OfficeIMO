"""Aggregates results.csv and visual.csv into scorecard tables and worst-case examples (scorecard.md)."""
from __future__ import annotations

import csv
import json
import math
import os
from collections import defaultdict
from difflib import SequenceMatcher

GROUPS = [
    ("Text and reading order", ["char_recall", "char_precision", "word_recall", "word_precision", "reading_order"]),
    ("Headings and lists", ["heading_recall", "heading_precision", "heading_level_acc", "list_recall", "list_precision",
                            "list_level_acc", "list_type_acc"]),
    ("Tables", ["table_recall", "table_precision", "teds", "teds_struct", "teds_all", "table_shape_acc", "merged_cells",
                "empty_cells"]),
    ("Objects and styles", ["images", "links", "footnote_text", "footnote_native", "slide_count", "bold_f1", "italic_f1",
                            "font_size_acc"]),
    ("Spreadsheet cells", ["xl_number_value", "xl_number_typed", "xl_date_value", "xl_date_typed", "xl_text"]),
]
TARGET_ORDER = ["docx", "odt", "rtf", "xlsx", "ods", "pptx", "odp", "html", "md"]
TOOL_ORDER = ["officeimo", "libreoffice", "pdf2docx", "docling", "pymupdf4llm", "pdfoxide"]


def _read(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return math.nan


def _mean(values):
    vals = [v for v in values if not math.isnan(v)]
    return (sum(vals) / len(vals), len(vals)) if vals else (math.nan, 0)


def _fmt(v, n=None):
    if math.isnan(v):
        return "–"
    return f"{v:.2f}" + (f" ({n})" if n is not None else "")


def scorecard(rows, visual, filt=lambda r: True):
    """{(target, tool): {metric: (mean, n)}}"""
    groups = defaultdict(list)
    for r in rows:
        if filt(r):
            groups[(r["target"], r["tool"])].append(r)
    vis = defaultdict(list)
    for v in visual:
        vis[(v["target"], v["tool"])].append(_num(v["visual_layout"]))
    out = {}
    for key, items in groups.items():
        m = {"docs": (len(items), len(items)),
             "success": (sum(r["status"] == "ok" for r in items) / len(items), len(items))}
        for _, metrics in GROUPS:
            for metric in metrics:
                m[metric] = _mean([_num(r.get(metric)) for r in items])
        m["visual_layout"] = _mean(vis.get(key, []))
        out[key] = m
    return out


def _table(card, metrics, title):
    lines = [f"#### {title}", "", "| Target | Tool | " + " | ".join(metrics) + " |",
             "|---|---|" + "---:|" * len(metrics)]
    for target in TARGET_ORDER:
        for tool in TOOL_ORDER:
            m = card.get((target, tool))
            if not m:
                continue
            cells = []
            for metric in metrics:
                mean, n = m.get(metric, (math.nan, 0))
                cells.append(_fmt(mean) if metric != "docs" else str(n))
            if all(c == "–" for c in cells) and metrics[0] != "docs":
                continue
            lines.append(f"| {target} | {tool} | " + " | ".join(cells) + " |")
    return "\n".join(lines) + "\n"


def _snippet(gold_words, out_words, width=14):
    sm = SequenceMatcher(None, gold_words, out_words, autojunk=False)
    worst = None
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        size = max(i2 - i1, j2 - j1)
        if worst is None or size > worst[0]:
            worst = (size, i1, i2, j1, j2)
    if worst is None:
        return None
    _, i1, i2, j1, j2 = worst
    g = " ".join(gold_words[max(0, i1 - 4):min(i2, i1 + width) + 2])
    o = " ".join(out_words[max(0, j1 - 4):min(j2, j1 + width) + 2])
    return g, o


def write_report(work: str) -> None:
    import sys

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from run import NORMALIZE, load_manifest, outputs_for, read_gold, read_output

    rows = _read(os.path.join(work, "results.csv"))
    visual = _read(os.path.join(work, "visual.csv"))
    manifest = load_manifest(work)
    lines = ["# PDF-to-X scorecard (generated)", ""]
    sets = defaultdict(lambda: [0, "", set()])
    for e in manifest:
        s = sets[e["set"]]
        s[0] += 1
        s[1] = e["licence"]
        s[2].add(e["family"])
    lines += ["## Corpus", "", "| Set | Documents | Families | Licence |", "|---|---:|---|---|"]
    for name, (count, licence, fams) in sorted(sets.items()):
        lines.append(f"| {name} | {count} | {', '.join(sorted(fams))} | {licence} |")
    lines.append("")
    def primary(r):
        # Spreadsheet targets extract tables only, so prose sources are scored for them in a separate table.
        return not (r["target"] in ("xlsx", "ods") and r["family"] != "excel")

    family = {e["id"]: e["family"] for e in manifest}
    card = scorecard([r for r in rows if primary(r)],
                     [v for v in visual if primary({"target": v["target"], "family": family.get(v["id"], "")})])
    lines += ["## Scorecard", "",
              "Means over documents where the metric applies; failed conversions score 0 on recall metrics. "
              "XLSX and ODS rows cover spreadsheet sources; tables from prose sources follow separately.", ""]
    lines.append(_table(card, ["docs", "success", "visual_layout"], "Coverage and visual similarity"))
    for title, metrics in GROUPS:
        lines.append(_table(card, metrics, title))
    tables = scorecard([r for r in rows if not primary(r)], [])
    lines.append(_table(tables, ["docs", "table_recall", "table_precision", "teds", "teds_struct", "teds_all",
                                 "table_shape_acc", "merged_cells", "empty_cells"], "Tables from prose sources to spreadsheets"))
    for fam in ("word", "tagged", "excel", "slides"):
        sub = scorecard([r for r in rows if r["family"] == fam], [])
        lines.append(f"### Family: {fam}\n")
        lines.append(_table(sub, ["docs", "success", "word_recall", "word_precision", "reading_order", "heading_recall",
                                  "list_recall", "table_recall", "teds"], f"{fam} summary"))
    # Worst cases for OfficeIMO
    lines += ["## Worst cases (OfficeIMO)", ""]
    by_id = {e["id"]: e for e in manifest}
    for target in TARGET_ORDER:
        cands = [r for r in rows if r["tool"] == "officeimo" and r["target"] == target and r["status"] == "ok"
                 and not math.isnan(_num(r["word_recall"])) and primary(r)]
        if not cands:
            continue
        cands.sort(key=lambda r: (_num(r["word_recall"]) if not math.isnan(_num(r["word_recall"])) else 0)
                   + (_num(r["reading_order"]) if not math.isnan(_num(r["reading_order"])) else 0))
        lines.append(f"### {target}\n")
        for r in cands[:4]:
            e = by_id[r["id"]]
            lines.append(f"- `{r['id']}` status {r['status']}, word recall {_fmt(_num(r['word_recall']))}, "
                         f"precision {_fmt(_num(r['word_precision']))}, order {_fmt(_num(r['reading_order']))}, "
                         f"TEDS {_fmt(_num(r['teds']))}. Source `{e['source']}`, PDF `{e['pdf']}`.")
            path = next((p for tool, t, p, _ in outputs_for(work, e) if tool == "officeimo" and t == target), None)
            if path:
                lines.append(f"  Output `{path}`.")
                try:
                    gold = read_gold(e).reading_words()
                    out = read_output(path, NORMALIZE.get(target, target)).reading_words(True)
                    snip = _snippet(gold, out)
                    if snip:
                        lines.append(f"  - source: `{snip[0][:220]}`")
                        lines.append(f"  - output: `{snip[1][:220]}`")
                except Exception as exc:
                    lines.append(f"  - (snippet unavailable: {type(exc).__name__})")
        lines.append("")
    with open(os.path.join(work, "scorecard.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    with open(os.path.join(work, "scorecard.json"), "w", encoding="utf-8") as f:
        json.dump({f"{k[0]}|{k[1]}": {m: v[0] for m, v in d.items()} for k, d in card.items()}, f, indent=1, default=str)
    print("wrote scorecard.md")
