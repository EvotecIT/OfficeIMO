"""Fidelity metrics between a ground-truth document and a converted document. All scores are in [0, 1]."""
from __future__ import annotations

import math
from bisect import bisect_left
from collections import Counter
from difflib import SequenceMatcher

from .model import Cell, Document, Table, norm_text, tokens
from .sheets import cell_value, parse_date, parse_number, parse_temporal_text, temporal_key, temporal_keys

NA = float("nan")


def _ratio(num: float, den: float) -> float:
    return num / den if den else NA


def _f1(p: float, r: float) -> float:
    if math.isnan(p) or math.isnan(r):
        return NA
    return 2 * p * r / (p + r) if p + r else 0.0


def levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        current = [i]
        for j, cb in enumerate(b, 1):
            current.append(min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (ca != cb)))
        previous = current
    return previous[-1]


def text_similarity(a: str, b: str) -> float:
    a, b = norm_text(a).casefold(), norm_text(b).casefold()
    if not a and not b:
        return 1.0
    if len(a) > 300 or len(b) > 300:
        return SequenceMatcher(None, a, b, autojunk=False).ratio()
    return 1 - levenshtein(a, b) / max(len(a), len(b))


# ------------------------------------------------------------------------------------------------ text

def text_metrics(gold: Document, out: Document) -> dict:
    gw = gold.reading_words()
    ow = out.reading_words(include_optional=True)
    g, o = Counter(gw), Counter(ow)
    # Header and footer text repeats once per page in a faithful output, so any number of copies is allowed.
    optional = set(gold.optional_words())
    allowed = Counter(g)
    for word in optional:
        allowed[word] = max(allowed[word], o[word])
    common = sum((g & o).values())
    word_r = _ratio(common, len(gw))
    word_p = _ratio(sum((o & allowed).values()), len(ow)) if gw else NA
    gc, oc = Counter("".join(gw)), Counter("".join(ow))
    ac = Counter("".join(allowed.elements()))
    char_r = _ratio(sum((gc & oc).values()), sum(gc.values()))
    char_p = _ratio(sum((oc & ac).values()), sum(oc.values())) if gw else NA
    # Reading order: share of recovered words that also keep the source order.
    order = NA
    limit = 40000
    shared = sum((Counter(gw[:limit]) & Counter(ow[:limit])).values())
    if shared:
        matcher = SequenceMatcher(None, gw[:limit], ow[:limit], autojunk=False)
        ordered = sum(block.size for block in matcher.get_matching_blocks())
        order = min(1.0, ordered / shared)
    return {"char_recall": char_r, "char_precision": char_p, "word_recall": word_r, "word_precision": word_p,
            "word_f1": _f1(word_p, word_r), "reading_order": order}


# ------------------------------------------------------------------------------------------------ headings and lists

def _match_blocks(gold_blocks, out_blocks, threshold=0.8):
    pairs = []
    used = set()
    index: dict[str, list[int]] = {}
    for j, b in enumerate(out_blocks):
        index.setdefault(" ".join(tokens(b.text)).casefold(), []).append(j)
    for gb in gold_blocks:
        key = " ".join(tokens(gb.text)).casefold()
        best, best_j = 0.0, None
        for j in index.get(key, []):
            if j not in used:
                best, best_j = 1.0, j
                break
        if best_j is None:
            for j, ob in enumerate(out_blocks):
                if j in used:
                    continue
                s = text_similarity(key, " ".join(tokens(ob.text)))
                if s > best:
                    best, best_j = s, j
        if best_j is not None and best >= threshold:
            used.add(best_j)
            pairs.append((gb, out_blocks[best_j]))
    return pairs


def _pairwise_levels(pairs) -> float:
    """Share of matched heading pairs whose relative level (higher, same, lower) agrees with the source.

    Independent of absolute numbering, so a document title emitted as H1 costs only the pairs it takes part in.
    With a single matched heading there is no relation to compare.
    """
    if len(pairs) < 2:
        return NA
    agree = total = 0
    sign = lambda x: (x > 0) - (x < 0)  # noqa: E731
    for i in range(len(pairs)):
        for j in range(i + 1, min(len(pairs), i + 60)):
            total += 1
            agree += sign(pairs[i][0].level - pairs[j][0].level) == sign(pairs[i][1].level - pairs[j][1].level)
    return agree / total


def structure_metrics(gold: Document, out: Document) -> dict:
    res = {}
    gh = [b for b in gold.blocks if b.kind == "heading" and not b.optional and tokens(b.text)]
    oh = [b for b in out.blocks if b.kind == "heading" and tokens(b.text)]
    pairs = _match_blocks(gh, oh)
    res["heading_recall"] = _ratio(len(pairs), len(gh))
    res["heading_precision"] = _ratio(len(pairs), len(oh)) if gh else NA
    res["heading_level_acc"] = _pairwise_levels(pairs)
    gl = [b for b in gold.blocks if b.kind == "list" and not b.optional and tokens(b.text)]
    ol = [b for b in out.blocks if b.kind == "list" and tokens(b.text)]
    pairs = _match_blocks(gl, ol)
    res["list_recall"] = _ratio(len(pairs), len(gl))
    res["list_precision"] = _ratio(len(pairs), len(ol)) if gl else NA
    if pairs:
        g0 = min(b.level for b in gl)
        o0 = min(b.level for b in ol)
        res["list_level_acc"] = sum((a.level - g0) == (b.level - o0) for a, b in pairs) / len(pairs)
        typed = [(a, b) for a, b in pairs if a.ordered is not None]
        res["list_type_acc"] = (sum(a.ordered == bool(b.ordered) for a, b in typed) / len(typed)) if typed else NA
    else:
        res["list_level_acc"] = res["list_type_acc"] = NA
    return res


# ------------------------------------------------------------------------------------------------ tables

def cell_cost(a: Cell, b: Cell, structure_only: bool = False) -> float:
    if (a.row_span, a.col_span) != (b.row_span, b.col_span):
        return 1.0
    if structure_only:
        return 0.0
    va, vb = cell_value(a), cell_value(b)
    if va is not None and vb is not None:
        tol = max(va[1], vb[1], 1e-9 * abs(va[0]))
        return 0.0 if abs(va[0] - vb[0]) <= tol + 1e-12 else 1.0
    da, db = parse_date(a.text), parse_date(b.text)
    if da and db:
        return 0.0 if set(da) & set(db) else 1.0
    return 1.0 - text_similarity(a.text, b.text)


def table_tree_similarity(g: Table, o: Table, structure_only: bool = False) -> float:
    """TEDS-style score: tree edit distance on table/row/cell trees with rows aligned to rows.

    Constraining rows to map onto rows makes this an upper bound on the unconstrained tree edit distance,
    which keeps large tables tractable. Node costs: insert/delete 1, cell rename = span mismatch or text distance.
    """
    def row_cost(gr, orow):
        n, m = len(gr), len(orow)
        prev = list(range(m + 1))
        for i in range(1, n + 1):
            cur = [i] + [0] * m
            for j in range(1, m + 1):
                cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cell_cost(gr[i - 1], orow[j - 1], structure_only))
            prev = cur
        return prev[m]

    rows_g, rows_o = g.rows[:300], o.rows[:300]
    n, m = len(rows_g), len(rows_o)
    prev = [0.0] * (m + 1)
    for j in range(1, m + 1):
        prev[j] = prev[j - 1] + 1 + len(rows_o[j - 1])
    for i in range(1, n + 1):
        cur = [prev[0] + 1 + len(rows_g[i - 1])] + [0.0] * m
        for j in range(1, m + 1):
            cur[j] = min(prev[j] + 1 + len(rows_g[i - 1]), cur[j - 1] + 1 + len(rows_o[j - 1]),
                         prev[j - 1] + row_cost(rows_g[i - 1], rows_o[j - 1]))
        prev = cur
    size = max(1 + n + sum(map(len, rows_g)), 1 + m + sum(map(len, rows_o)))
    return max(0.0, 1 - prev[m] / size)


def _grid(t: Table):
    grid: dict[tuple[int, int], Cell] = {}
    for r, row in enumerate(t.rows):
        c = 0
        for cell in row:
            while (r, c) in grid:
                c += 1
            for dr in range(cell.row_span):
                for dc in range(cell.col_span):
                    grid[(r + dr, c + dc)] = cell
            c += cell.col_span
    return grid


def _shape(t: Table):
    g = _grid(t)
    if not g:
        return 0, 0
    return max(r for r, _ in g) + 1, max(c for _, c in g) + 1


def _cells(t: Table) -> int:
    return sum(len(r) for r in t.rows)


class _CellBag:
    """Non-empty cells of a table for pairing: numbers by value within displayed precision (as `cell_cost` scores
    them), dates by serial, everything else by normalised text."""

    def __init__(self, table: Table):
        self.keys: Counter = Counter()
        numbers = []
        self.dates: list[set[float]] = []
        for row in table.rows:
            for c in row:
                if not norm_text(c.text):
                    continue
                value = cell_value(c)
                if value is not None:
                    numbers.append(value)
                    continue
                dates = parse_date(c.text)
                if dates:
                    self.dates.append(set(dates))  # every reading of an ambiguous date, as cell_cost accepts
                else:
                    self.keys[norm_text(c.text).casefold()] += 1
        numbers.sort()
        self.values = [v for v, _ in numbers]
        self.tols = [t for _, t in numbers]
        self.max_tol = max(self.tols, default=0.0)
        self.size = sum(self.keys.values()) + len(numbers) + len(self.dates)

    def overlap(self, other: "_CellBag") -> int:
        """Size of a one-to-one matching between the two bags."""
        inter = sum((self.keys & other.keys).values())
        used = [False] * len(other.values)
        for value, tol in zip(self.values, self.tols):
            reach = max(tol, other.max_tol, 1e-9 * abs(value)) + 1e-12
            i = bisect_left(other.values, value - reach)
            while i < len(other.values) and other.values[i] <= value + reach:
                if not used[i] and abs(other.values[i] - value) <= max(tol, other.tols[i], 1e-9 * abs(value)) + 1e-12:
                    used[i] = True
                    inter += 1
                    break
                i += 1
        taken = [False] * len(other.dates)
        for candidates in self.dates:
            k = next((k for k, o in enumerate(other.dates) if not taken[k] and candidates & o), None)
            if k is not None:
                taken[k] = True
                inter += 1
        return inter


def table_metrics(gold: Document, out: Document) -> dict:
    gts = [t for t in gold.tables if _cells(t) >= 2]
    ots = [t for t in out.tables if _cells(t) >= 2]
    res = {}
    if not gts:
        return {"table_recall": NA, "table_precision": NA, "teds": NA, "teds_struct": NA, "teds_all": NA,
                "table_shape_acc": NA, "merged_cells": NA, "empty_cells": NA}
    used = set()
    pairs = []
    obags = [_CellBag(ot) for ot in ots]
    for gt in gts:
        gbag = _CellBag(gt)
        if not gbag.size:
            # A table of empty cells can only be paired by its grid.
            j = next((j for j, ot in enumerate(ots) if j not in used and _shape(ot) == _shape(gt)
                      and not any(norm_text(c.text) for row in ot.rows for c in row)), None)
            if j is not None:
                used.add(j)
                pairs.append((gt, ots[j]))
            continue
        best, best_j = 0.0, None
        for j, ot in enumerate(ots):
            if j in used:
                continue
            obag = obags[j]
            inter = gbag.overlap(obag)
            if not inter:
                continue
            p, r = inter / max(1, obag.size), inter / max(1, gbag.size)
            f = 2 * p * r / (p + r)
            if f > best:
                best, best_j = f, j
        if best_j is not None and best >= 0.3:
            used.add(best_j)
            pairs.append((gt, ots[best_j]))
    res["table_recall"] = len(pairs) / len(gts)
    res["table_precision"] = _ratio(len(pairs), len(ots))
    teds = [table_tree_similarity(a, b) for a, b in pairs]
    res["teds"] = sum(teds) / len(teds) if teds else NA
    res["teds_all"] = sum(teds) / len(gts)
    res["teds_struct"] = (sum(table_tree_similarity(a, b, True) for a, b in pairs) / len(pairs)) if pairs else NA
    res["table_shape_acc"] = (sum(_shape(a) == _shape(b) for a, b in pairs) / len(pairs)) if pairs else NA
    merged_total = merged_ok = 0
    empty_total = empty_ok = 0
    for a, b in pairs:
        out_spans = Counter((norm_text(c.text).casefold(), c.row_span, c.col_span) for row in b.rows for c in row)
        for row in a.rows:
            for c in row:
                if (c.row_span > 1 or c.col_span > 1) and norm_text(c.text):
                    merged_total += 1
                    key = (norm_text(c.text).casefold(), c.row_span, c.col_span)
                    if out_spans[key] > 0:
                        merged_ok += 1
                        out_spans[key] -= 1
        ga, gb = _grid(a), _grid(b)
        for pos, cell in ga.items():
            if not norm_text(cell.text):
                empty_total += 1
                other = gb.get(pos)
                if other is not None and not norm_text(other.text):
                    empty_ok += 1
    res["merged_cells"] = _ratio(merged_ok, merged_total)
    res["empty_cells"] = _ratio(empty_ok, empty_total)
    return res


# ------------------------------------------------------------------------------------------------ images, links, notes

def _norm_url(u: str) -> str:
    """Scheme and host are case-insensitive; path and query are compared exactly."""
    from urllib.parse import urlsplit, urlunsplit

    parts = urlsplit(u.strip())
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), parts.path.rstrip("/"), parts.query, parts.fragment))


def object_metrics(gold: Document, out: Document) -> dict:
    res = {"images": _ratio(min(out.images, gold.images), gold.images) if gold.images else NA}
    gl = {_norm_url(u) for u in gold.links}
    ol = {_norm_url(u) for u in out.links}
    res["links"] = _ratio(len(gl & ol), len(gl)) if gl else NA
    if gold.footnotes:
        out_bag = Counter(out.reading_words(include_optional=True))
        text_hits = native_hits = 0
        native_used: set[int] = set()
        for note in gold.footnotes:
            nt = Counter(tokens(note))
            if nt and sum((nt & out_bag).values()) / sum(nt.values()) >= 0.8:
                text_hits += 1
            if out.footnotes_native:
                k = next((k for k, o in enumerate(out.footnotes) if k not in native_used and text_similarity(note, o) >= 0.8), None)
                if k is not None:
                    native_used.add(k)
                    native_hits += 1
        res["footnote_text"] = text_hits / len(gold.footnotes)
        res["footnote_native"] = native_hits / len(gold.footnotes)
    else:
        res["footnote_text"] = res["footnote_native"] = NA
    if gold.pages and out.pages is not None:
        res["slide_count"] = 1.0 if gold.pages == out.pages else max(0.0, 1 - abs(gold.pages - out.pages) / gold.pages)
    else:
        res["slide_count"] = NA
    return res


# ------------------------------------------------------------------------------------------------ styles

def style_metrics(gold: Document, out: Document) -> dict:
    gw = list(gold.styled_words())
    ow = list(out.styled_words())
    res = {"bold_f1": NA, "italic_f1": NA, "font_size_acc": NA}
    if not gw or not ow:
        return res
    gt = [tokens(w.text)[0] if tokens(w.text) else "" for w in gw]
    ot = [tokens(w.text)[0] if tokens(w.text) else "" for w in ow]
    matcher = SequenceMatcher(None, gt[:30000], ot[:30000], autojunk=False)
    pairs = []
    for block in matcher.get_matching_blocks():
        for k in range(block.size):
            if gt[block.a + k]:
                pairs.append((gw[block.a + k], ow[block.b + k]))
    for attr, key in (("bold", "bold_f1"), ("italic", "italic_f1")):
        tp = sum(getattr(a, attr) and getattr(b, attr) for a, b in pairs)
        fn = sum(getattr(a, attr) and not getattr(b, attr) for a, b in pairs)
        fp = sum(not getattr(a, attr) and getattr(b, attr) for a, b in pairs)
        if tp + fn:
            res[key] = 2 * tp / (2 * tp + fn + fp)
    sized = [(a.size, b.size) for a, b in pairs if a.size and b.size]
    if sized:
        res["font_size_acc"] = sum(abs(a - b) <= 0.6 for a, b in sized) / len(sized)
    return res


# ------------------------------------------------------------------------------------------------ spreadsheets

def sheet_metrics(gold: Document, out: Document) -> dict:
    res = {"xl_number_value": NA, "xl_number_typed": NA, "xl_date_value": NA, "xl_date_typed": NA, "xl_text": NA}
    if not gold.cells:
        return res
    out_cells = out.cells
    if not out_cells:
        # Non-spreadsheet outputs: compare against table cell text.
        from .model import SheetCell
        out_cells = [SheetCell(0, 0, c.text, "text") for t in out.tables for row in t.rows for c in row if c.text]
        out_cells += [SheetCell(0, 0, b.text, "text") for b in out.blocks]
    numbers = [c for c in gold.cells if c.kind == "number"]
    dates = [c for c in gold.cells if c.kind == "date" and c.value is not None]
    texts = [c for c in gold.cells if c.kind == "text"]
    parsed = []
    for c in out_cells:
        typed = c.kind == "number" and c.value is not None
        value = (c.value, 0.0) if typed else parse_number(c.text)
        parsed.append((typed, value, c))
    used: set[int] = set()

    def find_number(g):
        tol_g = 0.5 * 10 ** (-(g.decimals if g.decimals is not None else 9)) / (100 if g.percent else 1)
        hit = None
        for i, (typed, value, _) in enumerate(parsed):
            if i in used or value is None:
                continue
            if abs(value[0] - g.value) <= max(tol_g, value[1], 1e-9 * abs(g.value)) + 1e-12:
                if typed:
                    used.add(i)
                    return "typed"
                if hit is None:
                    hit = i
        if hit is not None:
            used.add(hit)
            return "text"
        return None

    if numbers:
        found = [find_number(g) for g in numbers[:4000]]
        res["xl_number_value"] = sum(f is not None for f in found) / len(found)
        res["xl_number_typed"] = sum(f == "typed" for f in found) / len(found)
    if dates:
        typed_hits = any_hits = 0
        # Each output cell answers for one source date: typed date serials first, then dates read from text.
        # A source value with a time of day must keep that time (to the minute); a plain date compares by day.
        typed_keys = [temporal_keys(c.value) if c.kind == "date" and c.value is not None else set() for c in out_cells]
        text_keys = [parse_temporal_text(c.text) for c in out_cells]
        used_dates: set[int] = set()
        for g in dates[:2000]:
            key = temporal_key(g.value)
            hit = next((i for i, s in enumerate(typed_keys) if key in s and i not in used_dates), None)
            typed = hit is not None
            if hit is None:
                hit = next((i for i, s in enumerate(text_keys) if key in s and i not in used_dates), None)
            if hit is not None:
                used_dates.add(hit)
            typed_hits += typed
            any_hits += hit is not None
        res["xl_date_typed"] = typed_hits / len(dates[:2000])
        res["xl_date_value"] = any_hits / len(dates[:2000])
    if texts:
        bag = Counter(norm_text(c.text).casefold() for c in out_cells)
        hits = 0
        for g in texts[:4000]:
            key = norm_text(g.text).casefold()
            if bag[key] > 0:
                bag[key] -= 1
                hits += 1
        res["xl_text"] = hits / len(texts[:4000])
    return res


def score(gold: Document, out: Document) -> dict:
    res = {}
    for fn in (text_metrics, structure_metrics, table_metrics, object_metrics, style_metrics, sheet_metrics):
        res.update(fn(gold, out))
    return res
