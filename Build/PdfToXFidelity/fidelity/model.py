"""Format-neutral document model shared by the ground-truth and output readers."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field


@dataclass
class Word:
    text: str
    bold: bool = False
    italic: bool = False
    size: float | None = None


@dataclass
class Block:
    kind: str  # heading | paragraph | list | caption
    words: list[Word] = field(default_factory=list)
    level: int = 0  # heading level or list nesting level (0-based for lists)
    ordered: bool = False
    optional: bool = False  # header/footer text: neither required nor penalised

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)


@dataclass
class Cell:
    text: str
    row_span: int = 1
    col_span: int = 1
    value: float | None = None  # typed spreadsheet number, compared numerically
    tol: float = 0.0


@dataclass
class Table:
    rows: list[list[Cell]] = field(default_factory=list)

    def texts(self) -> list[str]:
        return [norm_text(c.text) for r in self.rows for c in r]


@dataclass
class SheetCell:
    row: int
    col: int
    text: str
    kind: str  # number | date | bool | text | error
    value: float | None = None
    decimals: int | None = None  # displayed decimals when the number format states them
    percent: bool = False


@dataclass
class Document:
    """Reading-ordered content. Tables also appear as `table` markers in `order` so reading order includes them."""

    blocks: list[Block] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    order: list[tuple[str, int]] = field(default_factory=list)  # ("block", i) | ("table", i)
    images: int = 0
    links: set[str] = field(default_factory=set)
    footnotes: list[str] = field(default_factory=list)
    footnotes_native: bool = False  # footnotes carried as real notes rather than plain text
    cells: list[SheetCell] = field(default_factory=list)
    pages: int | None = None  # slides for presentations
    notes: list[str] = field(default_factory=list)

    def add_block(self, block: Block) -> None:
        if not block.words:
            return
        self.order.append(("block", len(self.blocks)))
        self.blocks.append(block)

    def add_table(self, table: Table) -> None:
        if not table.rows:
            return
        self.order.append(("table", len(self.tables)))
        self.tables.append(table)

    def reading_words(self, include_optional: bool = False) -> list[str]:
        """All text in reading order, tables row-major, footnotes last."""
        out: list[str] = []
        for kind, index in self.order:
            if kind == "block":
                b = self.blocks[index]
                if b.optional and not include_optional:
                    continue
                out.extend(tokens(b.text))
            else:
                for row in self.tables[index].rows:
                    for cell in row:
                        out.extend(tokens(cell.text))
        for note in self.footnotes:
            out.extend(tokens(note))
        return out

    def optional_words(self) -> list[str]:
        out: list[str] = []
        for b in self.blocks:
            if b.optional:
                out.extend(tokens(b.text))
        return out

    def styled_words(self):
        for b in self.blocks:
            if not b.optional:
                yield from b.words


_space = re.compile(r"\s+")
_token = re.compile(r"\w+(?:[.,'’]\w+)*|[^\w\s]", re.UNICODE)
_dash = str.maketrans({"‐": "-", "‑": "-", "‒": "-", "–": "-", "—": "-", "−": "-",
                       "‘": "'", "’": "'", "“": '"', "”": '"', " ": " ", " ": " ",
                       " ": " ", "­": "", "​": "", "﻿": ""})


def norm_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "").translate(_dash)
    return _space.sub(" ", value).strip()


_numeric = re.compile(r"^\d[\d.,']*$")


def canonical_number(token: str) -> str:
    """Numbers compare by value, so 1,234.50 / 1.234,50 / 1234.5 are the same token."""
    if not _numeric.match(token) or not any(ch in token for ch in ".,'"):
        return (token.lstrip("0") or "0") if token.isdigit() else token
    from .sheets import parse_number

    parsed = parse_number(token)
    if parsed is None:
        return token
    return format(round(parsed[0], 6), "g") if abs(parsed[0]) < 1e15 else token


def tokens(value: str) -> list[str]:
    """Word tokens with punctuation dropped and numbers canonicalised; bullets and list markers carry no content."""
    return [canonical_number(t) for t in _token.findall(norm_text(value)) if any(ch.isalnum() for ch in t)]


def split_words(text: str, bold=False, italic=False, size=None) -> list[Word]:
    return [Word(t, bold, italic, size) for t in norm_text(text).split(" ") if t]
