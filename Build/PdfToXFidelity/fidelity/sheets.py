"""Spreadsheet grid segmentation and culture-tolerant cell value parsing."""
from __future__ import annotations

import datetime as dt
import re

from .model import Block, Cell, Document, Table, norm_text, split_words

_DATE_PATTERNS = [
    (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$"), ("y", "m", "d")),
    (re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})\.?$"), ("d", "m", "y")),
    (re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$"), ("m", "d", "y")),
    (re.compile(r"^(\d{1,2})-(\d{1,2})-(\d{2,4})$"), ("d", "m", "y")),
    (re.compile(r"^(\d{4})/(\d{1,2})/(\d{1,2})$"), ("y", "m", "d")),
    (re.compile(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})\.?$"), ("y", "m", "d")),
]
_MONTHS = {m: i + 1 for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
_NAMED_DATE = re.compile(r"^(\d{1,2})[ .-]([A-Za-z]{3})[A-Za-z]*\.?[ .-](\d{2,4})$")
_EPOCH = dt.date(1899, 12, 30)
MAX_TABLE_COLUMNS = 61


def parse_date(text: str) -> list[float]:
    """Candidate Excel serials for a displayed date; slash dates are ambiguous, so both readings are returned."""
    value = norm_text(text)
    out = []
    m = _NAMED_DATE.match(value)
    if m and m.group(2).lower() in _MONTHS:
        year = _full_year(int(m.group(3)))
        try:
            out.append(float((dt.date(year, _MONTHS[m.group(2).lower()], int(m.group(1))) - _EPOCH).days))
        except ValueError:
            pass
    for pattern, order in _DATE_PATTERNS:
        m = pattern.match(value)
        if not m:
            continue
        parts = dict(zip(order, (int(g) for g in m.groups()[:3])))
        readings = [parts]
        if order == ("m", "d", "y"):
            readings.append({"d": parts["m"], "m": parts["d"], "y": parts["y"]})
        for p in readings:
            year = _full_year(p["y"])
            try:
                out.append(float((dt.date(year, p["m"], p["d"]) - _EPOCH).days))
            except ValueError:
                pass
    return out


def _full_year(year: int) -> int:
    """Two-digit years follow the spreadsheet rule: 00-49 are 20xx, 50-99 are 19xx."""
    return year + (2000 if year < 50 else 1900 if year < 100 else 0)


_SCIENTIFIC = re.compile(r"(\d+)(?:[.,](\d+))?[eE]([+-]?\d+)")
_CURRENCY = re.compile(r"[$€£¥₹]|zł|kr|chf|eur|usd|pln|gbp|sek", re.I)


def parse_number(text: str) -> tuple[float, float] | None:
    """(value, tolerance) for a displayed number in en, de, fr, pl or ch style; None when the text is not a number."""
    s = norm_text(text)
    if not s or parse_date(s) or len(s) > 40:
        return None
    s = _CURRENCY.sub("", s).strip()
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1].strip()
    percent = s.endswith("%")
    if percent:
        s = s[:-1].strip()
    if s.startswith(("-", "+")):
        negative, s = negative or s[0] == "-", s[1:].strip()
    if s.endswith("-"):
        negative, s = True, s[:-1].strip()
    s = s.replace(" ", "").replace("'", "")
    sci = _SCIENTIFIC.fullmatch(s)
    if sci:
        # 1.23E-05: the mantissa's decimals set the displayed precision, scaled by the exponent.
        scale = 10.0 ** int(sci.group(3))
        frac = sci.group(2) or ""
        value, tol = float(sci.group(1) + "." + (frac or "0")) * scale, 0.5 * 10 ** (-len(frac)) * scale
        if percent:
            value, tol = value / 100, tol / 100
        return (-value if negative else value), tol
    if not re.fullmatch(r"\d[\d.,]*", s):
        return None
    decimals = 0
    if "." in s and "," in s:
        dec = "." if s.rfind(".") > s.rfind(",") else ","
        thousands = "," if dec == "." else "."
        int_part, frac = s.rsplit(dec, 1)
        s, decimals = int_part.replace(thousands, "") + "." + frac, len(frac)
    elif "," in s or "." in s:
        sep = "," if "," in s else "."
        pieces = s.split(sep)
        if len(pieces) > 2 or (len(pieces[1]) == 3 and pieces[0] not in ("0", "") and len(pieces[0]) <= 3 and sep == ","):
            s = s.replace(sep, "")
        elif len(pieces) == 2 and len(pieces[1]) == 3 and sep == "." and pieces[0] not in ("0", "") and len(pieces[0]) <= 3:
            s = s.replace(sep, "")  # 1.234 reads as German thousands; recorded as a known ambiguity
        else:
            s, decimals = pieces[0] + "." + pieces[1], len(pieces[1])
    try:
        value = float(s)
    except ValueError:
        return None
    tol = 0.5 * 10 ** (-decimals)
    if percent:
        value, tol = value / 100, tol / 100
    return (-value if negative else value), tol


def cell_value(cell: Cell) -> tuple[float, float] | None:
    if cell.value is not None:
        return cell.value, cell.tol
    return parse_number(cell.text)


def sheet_regions(doc: Document, cells: dict, merges: list) -> None:
    """Splits a sheet at fully empty rows; regions with two or more filled columns become tables."""
    if not cells:
        return
    covered = {}
    spans = {}
    for r1, c1, r2, c2 in merges:
        spans[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1):
                    covered[(r, c)] = (r1, c1)
    rows = sorted({r for r, _ in cells})
    groups: list[list[int]] = []
    for r in rows:
        if groups and r - groups[-1][-1] <= 1 + _merged_gap(groups[-1][-1], r, spans):
            groups[-1].append(r)
        else:
            groups.append([r])
    for group in groups:
        r0, r1 = group[0], group[-1]
        region = {k: v for k, v in cells.items() if r0 <= k[0] <= r1}
        width = max(sum(1 for (r, _) in region if r == row) for row in group)
        if width < 2:
            for row in group:
                for (r, c), cell in sorted(region.items()):
                    if r == row:
                        doc.add_block(Block("paragraph", split_words(cell.text)))
            continue
        c0 = min(c for _, c in region)
        c1 = max(c + spans.get(k, (1, 1))[1] - 1 for k in region for c in [k[1]])
        if c1 - c0 + 1 > MAX_TABLE_COLUMNS:
            # Too wide to score as a table: keep the full text row by row instead of a truncated grid.
            for row in group:
                words = [w for (r, _), cell in sorted(region.items()) if r == row for w in split_words(cell.text)]
                doc.add_block(Block("paragraph", words))
            continue
        table = Table()
        for r in range(r0, r1 + 1):
            row = []
            for c in range(c0, c1 + 1):
                if (r, c) in covered:
                    continue
                cell = region.get((r, c))
                rs, cs = spans.get((r, c), (1, 1))
                if cell is None:
                    row.append(Cell("", rs, cs))
                else:
                    value = cell.value
                    tol = 0.0
                    if cell.kind == "number" and value is not None:
                        tol = 0.5 * 10 ** (-(cell.decimals if cell.decimals is not None else 9))
                        if cell.percent:
                            tol /= 100
                    else:
                        value = None
                    row.append(Cell(norm_text(cell.text), rs, cs, value, tol))
            table.rows.append(row)
        doc.add_table(table)


def _merged_gap(last_row: int, next_row: int, spans: dict) -> int:
    # Rows still covered by a vertical merge after last_row do not split a region.
    return max((r + rs - 1 - last_row for (r, _), (rs, _) in spans.items() if r <= last_row < r + rs), default=0)
