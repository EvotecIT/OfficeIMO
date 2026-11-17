"""HTML, Markdown and tagged-PDF structure readers built on the standard library."""
from __future__ import annotations

import html
import re
from html.parser import HTMLParser

from .model import Block, Cell, Document, Table, Word, join_segments, norm_text, split_words

_BLOCK = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "br", "section", "article", "header", "footer", "aside",
          "blockquote", "pre", "caption", "figcaption", "dt", "dd", "tr", "td", "th", "table", "ul", "ol", "hr", "figure",
          "main", "nav", "body", "address", "summary", "details"}
_SKIP = {"script", "style", "head", "title", "template", "noscript", "svg", "math"}
_VOID = {"br", "img", "hr", "meta", "link", "input", "col", "source", "wbr", "area", "base", "embed"}


def _css_props(text: str) -> dict[str, str]:
    out = {}
    for decl in (text or "").split(";"):
        name, _, value = decl.partition(":")
        if value:
            out[name.strip().lower()] = value.strip().lower()
    return out


def _size_pt(value: str | None) -> float | None:
    if not value:
        return None
    m = re.match(r"([\d.]+)\s*(pt|px)", value)
    if not m:
        return None
    number = float(m.group(1))
    return number if m.group(2) == "pt" else number * 0.75


class _HtmlReader(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.doc = Document()
        self.classes: dict[str, dict[str, str]] = {}
        self.stack: list[tuple[str, dict]] = []  # (tag, style state)
        self.skip = 0
        self.words: list[Word] = []
        self.segments: list[tuple[str, bool, bool, float | None]] = []
        self.kind = ("paragraph", 0, False)
        self.lists: list[bool] = []
        self.tables: list[dict] = []
        self.footnote_depth = 0
        self.note_words: list[str] = []
        self.style_text: list[str] = []
        self.in_style = False

    # state helpers
    def _state(self) -> dict:
        return self.stack[-1][1] if self.stack else {"b": False, "i": False, "size": None}

    def _flush(self):
        self.words.extend(join_segments(self.segments))
        self.segments = []
        if self.tables:
            cell = self.tables[-1].get("cell")
            if cell is not None:
                cell.extend(self.words)
            self.words = []
            return
        kind, level, ordered = self.kind
        if self.footnote_depth:
            self.note_words.extend(w.text for w in self.words)
        elif self.words:
            self.doc.add_block(Block(kind, self.words, level, ordered))
        self.words = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "style":
            self.in_style = True
        if tag in _SKIP:
            if tag not in _VOID:
                self.skip += 1
            return
        if self.skip:
            return
        if tag == "img":
            self.doc.images += 1
            return
        if tag == "a" and a.get("href", "").lower().startswith(("http:", "https:", "mailto:", "ftp:")):
            self.doc.links.add(a["href"])
        if tag in _BLOCK:
            self._flush()
        state = dict(self._state())
        props = {}
        for cls in (a.get("class") or "").split():
            props.update(self.classes.get(cls, {}))
        props.update(_css_props(a.get("style", "")))
        weight = props.get("font-weight", "")
        numeric = int(weight) if weight.isdigit() else None
        # Headings and table headers render bold by default in browsers and Markdown viewers.
        if tag in ("b", "strong", "th", "h1", "h2", "h3", "h4", "h5", "h6") or weight in ("bold", "bolder") or (numeric is not None and numeric >= 600):
            state["b"] = True
        if weight in ("normal", "lighter") or (numeric is not None and numeric < 600):
            state["b"] = False
        if tag in ("i", "em") or props.get("font-style") == "italic":
            state["i"] = True
        size = _size_pt(props.get("font-size"))
        if size:
            state["size"] = size
        m = re.fullmatch(r"h([1-6])", tag)
        if m:
            self.kind = ("heading", int(m.group(1)), False)
        elif tag in ("ul", "ol"):
            self.lists.append(tag == "ol")
        elif tag == "li":
            self.kind = ("list", max(0, len(self.lists) - 1), bool(self.lists and self.lists[-1]))
        elif tag in ("p", "div") and not self.lists:
            self.kind = ("paragraph", 0, False)
        role = (a.get("role") or "") + " " + (a.get("class") or "") + " " + (a.get("id") or "")
        is_note = re.search(r"doc-footnote|doc-endnote|\bfootnotes?\b|^\s*fn\d", role) is not None
        if is_note and tag in ("li", "div", "section", "aside", "ol", "p"):
            if tag == "li" and self.footnote_depth:
                state["note_item"] = True  # an identified note inside a shared footnotes container
            self.footnote_depth += 1
            state["note"] = True
        elif tag == "li" and self.footnote_depth:
            state["note_item"] = True  # one note inside a shared footnotes list
        if tag == "table":
            self.tables.append({"table": Table(), "row": None, "cell": None})
        elif tag == "tr" and self.tables:
            self.tables[-1]["row"] = []
        elif tag in ("td", "th") and self.tables:
            t = self.tables[-1]
            if t["row"] is None:
                t["row"] = []
            t["cell"] = []
            t["span"] = (int(a.get("rowspan") or 1), int(a.get("colspan") or 1))
        if tag not in _VOID:
            self.stack.append((tag, state))

    def handle_endtag(self, tag):
        if tag == "style":
            self.in_style = False
            self._parse_classes("".join(self.style_text))
            self.style_text = []
        if tag in _SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip or tag in _VOID:
            return
        if tag in _BLOCK:
            self._flush()
        if tag in ("td", "th") and self.tables:
            t = self.tables[-1]
            if t["cell"] is not None and t["row"] is not None:
                rs, cs = t.get("span", (1, 1))
                t["row"].append(Cell(norm_text(" ".join(w.text for w in t["cell"])), rs, cs))
            t["cell"] = None
        elif tag == "tr" and self.tables:
            t = self.tables[-1]
            if t["row"] is not None:  # a row of rowspan continuations has no cells but still occupies a grid row
                t["table"].rows.append(t["row"])
            t["row"] = None
        elif tag == "table" and self.tables:
            t = self.tables.pop()
            if self.tables and self.tables[-1].get("cell") is not None:
                # Nested table text belongs to the enclosing cell.
                self.tables[-1]["cell"].extend(Word(x) for x in " ".join(" ".join(c.text for c in r) for r in t["table"].rows).split())
            else:
                self.doc.add_table(t["table"])
        elif tag in ("ul", "ol") and self.lists:
            self.lists.pop()
        # pop the matching element
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                if self.stack[i][1].get("note_item") and self.note_words:
                    self.doc.footnotes.append(" ".join(self.note_words))
                    self.doc.footnotes_native = True
                    self.note_words = []
                if self.stack[i][1].get("note") and not any(s.get("note") for _, s in self.stack[:i]):
                    self.footnote_depth = 0
                    if self.note_words:
                        self.doc.footnotes.append(" ".join(self.note_words))
                        self.doc.footnotes_native = True
                    self.note_words = []
                del self.stack[i:]
                break
        if re.fullmatch(r"h[1-6]|li|p|div", tag):
            self.kind = ("list", max(0, len(self.lists) - 1), self.lists[-1]) if self.lists and tag != "li" else ("paragraph", 0, False)

    def handle_data(self, data):
        if self.in_style:
            self.style_text.append(data)
            return
        if self.skip:
            return
        s = self._state()
        self.segments.append((data, s["b"], s["i"], s["size"]))

    def _parse_classes(self, css: str):
        for selector, body in re.findall(r"([^{}]+)\{([^}]*)\}", css):
            for sel in selector.split(","):
                m = re.fullmatch(r"\s*(?:[a-z0-9]*)\.([\w-]+)\s*", sel, re.I)
                if m:
                    self.classes.setdefault(m.group(1), {}).update(_css_props(body))

    def close(self):
        super().close()
        self._flush()


def read_html(path: str) -> Document:
    with open(path, encoding="utf-8", errors="replace") as f:
        return parse_html(f.read())


def parse_html(text: str) -> Document:
    reader = _HtmlReader()
    reader.feed(text)
    reader.close()
    return reader.doc


# ----------------------------------------------------------------------------------------------- Markdown

_INLINE_LINK = re.compile(r"!?\[([^\]]*)\]\(\s*<?([^)\s>]+)>?(?:\s+\"[^\"]*\")?\s*\)")
_AUTOLINK = re.compile(r"<((?:https?|mailto|ftp):[^>\s]+)>")
_TAG = re.compile(r"<[^>]+>")


def _md_inline(text: str, doc: Document) -> list[Word]:
    doc.images += len(re.findall(r"!\[[^\]]*\]\(", text)) + len(re.findall(r"<img\b", text, re.I))
    for m in _INLINE_LINK.finditer(text):
        if not m.group(0).startswith("!") and m.group(2).lower().startswith(("http:", "https:", "mailto:", "ftp:")):
            doc.links.add(m.group(2))
    for m in _AUTOLINK.finditer(text):
        doc.links.add(m.group(1))
    text = _INLINE_LINK.sub(lambda m: "" if m.group(0).startswith("!") else m.group(1), text)
    text = _AUTOLINK.sub(lambda m: m.group(1), text)
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = _TAG.sub("", text)
    text = re.sub(r"\[\^[^\]]+\]", "", text)  # footnote references
    text = html.unescape(text)
    segments: list[tuple[str, bool, bool, float | None]] = []
    bold = italic = False
    i = 0
    buf: list[str] = []

    def flush():
        segments.append((re.sub(r"\\(.)", r"\1", "".join(buf)), bold, italic, None))
        buf.clear()

    while i < len(text):
        ch = text[i]
        if ch == "\\" and i + 1 < len(text):
            buf.append(text[i:i + 2])
            i += 2
            continue
        if ch == "`":
            j = text.find("`", i + 1)
            if j > i:
                buf.append(text[i + 1:j])
                i = j + 1
                continue
        if ch in "*_":
            run = len(text) - len(text[i:].lstrip(ch))
            run = min(run, 3)
            prev = text[i - 1] if i else " "
            nxt = text[i + run] if i + run < len(text) else " "
            intraword_underscore = ch == "_" and prev.isalnum() and nxt.isalnum()
            if not intraword_underscore and not (prev.isspace() and nxt.isspace()):
                flush()
                if run >= 2:
                    bold = not bold
                if run in (1, 3):
                    italic = not italic
                i += run
                continue
        buf.append(ch)
        i += 1
    flush()
    return join_segments(segments)


_TABLE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


def _split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|") and not line.endswith("\\|"):
        line = line[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", line)]


def read_markdown(path: str) -> Document:
    with open(path, encoding="utf-8", errors="replace") as f:
        lines = f.read().replace("\r\n", "\n").split("\n")
    doc = Document()
    para: list[str] = []
    indents: list[int] = []
    i = 0

    def flush_para():
        if para:
            doc.add_block(Block("paragraph", _md_inline(" ".join(para), doc)))
            para.clear()

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("<!--"):
            while i < len(lines) and "-->" not in lines[i]:
                i += 1
            i += 1
            continue
        if stripped.startswith(("```", "~~~")):
            flush_para()
            fence = stripped[:3]
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith(fence):
                code.append(lines[i])
                i += 1
            doc.add_block(Block("paragraph", split_words(" ".join(code))))
            i += 1
            continue
        if re.match(r"\s*<table\b", line, re.I):
            flush_para()
            chunk = []
            while i < len(lines):
                chunk.append(lines[i])
                if re.search(r"</table>", lines[i], re.I):
                    break
                i += 1
            sub = parse_html("\n".join(chunk))
            for t in sub.tables:
                doc.add_table(t)
            doc.images += sub.images
            doc.links |= sub.links
            i += 1
            continue
        if not stripped:
            # A blank line does not end a loose list; content that leaves the list resets nesting below.
            flush_para()
            i += 1
            continue
        m = re.match(r"^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$", line)
        if m:
            flush_para()
            words = _md_inline(m.group(2), doc)
            for w in words:
                w.bold = True
            doc.add_block(Block("heading", words, len(m.group(1))))
            i += 1
            continue
        if para and re.match(r"^\s{0,3}(=+|-+)\s*$", line) and len(para) == 1:
            level = 1 if stripped[0] == "=" else 2
            words = _md_inline(para[0], doc)
            for w in words:
                w.bold = True  # rendered bold like ATX headings
            doc.add_block(Block("heading", words, level))
            para.clear()
            i += 1
            continue
        if "|" in line and i + 1 < len(lines) and _TABLE_SEP.match(lines[i + 1]):
            flush_para()
            table = Table([[Cell(norm_text(" ".join(w.text for w in _md_inline(c, doc)))) for c in _split_row(line)]])
            i += 2
            while i < len(lines) and "|" in lines[i] and lines[i].strip():
                table.rows.append([Cell(norm_text(" ".join(w.text for w in _md_inline(c, doc)))) for c in _split_row(lines[i])])
                i += 1
            doc.add_table(table)
            continue
        m = re.match(r"^\[\^([^\]]+)\]:\s*(.*)$", stripped)
        if m:
            flush_para()
            doc.footnotes.append(" ".join(w.text for w in _md_inline(m.group(2), doc)))
            doc.footnotes_native = True
            i += 1
            continue
        m = re.match(r"^(\s*)([-*+•]|\d{1,3}[.)])\s+(.*)$", line)
        if m and not re.match(r"^\s*([-*_])\s*(\1\s*){2,}$", line):
            flush_para()
            indent = len(m.group(1).replace("\t", "    "))
            while indents and indents[-1] > indent:
                indents.pop()
            if not indents or indents[-1] < indent:
                indents.append(indent)
            doc.add_block(Block("list", _md_inline(m.group(3), doc), len(indents) - 1, m.group(2)[0].isdigit()))
            i += 1
            continue
        if re.match(r"^\s{0,3}([-*_])\s*(\1\s*){2,}$", line):
            flush_para()
            i += 1
            continue
        if not line[:1].isspace():
            indents.clear()
        para.append(stripped.lstrip(">").strip())
        i += 1
    flush_para()
    return doc


# ----------------------------------------------------------------------------------------------- tagged PDF structure

_HEADING = re.compile(r"^H([1-6])?$")


def parse_struct_text(text: str) -> Document:
    """Ground truth from Poppler's `pdfinfo -struct-text`: logical structure order with element types."""
    root = {"type": "root", "children": [], "depth": -1}
    stack = [root]
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines):
        raw = lines[i]
        if not raw.strip():
            i += 1
            continue
        depth = (len(raw) - len(raw.lstrip(" "))) // 2
        body = raw.strip()
        if body.startswith('"'):
            value = body[1:]
            while not value.endswith('"') and i + 1 < len(lines):
                i += 1
                value += "\n" + lines[i]
            node = {"type": "#text", "text": value[:-1] if value.endswith('"') else value, "children": [], "depth": depth}
        else:
            node = {"type": body.split(" ")[0], "children": [], "depth": depth}
        while stack[-1]["depth"] >= depth:
            stack.pop()
        stack[-1]["children"].append(node)
        stack.append(node)
        i += 1
    doc = Document()

    def text_of(n) -> str:
        if n["type"] == "#text":
            return n["text"]
        return "".join(text_of(c) for c in n["children"])

    def own_text(n) -> str:
        """Text of a list item body without the items of lists nested inside it."""
        if n["type"] == "#text":
            return n["text"]
        return "".join(own_text(c) for c in n["children"] if c["type"] != "L")

    def nested_lists(n):
        for c in n["children"]:
            if c["type"] == "L":
                yield c
            elif c["type"] != "#text":
                yield from nested_lists(c)

    def walk(n, list_depth: int):
        t = n["type"]
        m = _HEADING.match(t)
        if m:
            doc.add_block(Block("heading", split_words(text_of(n)), int(m.group(1) or 1)))
        elif t in ("P", "Caption", "BlockQuote", "Note", "Code", "TOCI", "Lbl") and not any(c["type"] in ("P", "L", "Table") for c in n["children"]):
            doc.add_block(Block("paragraph", split_words(text_of(n))))
        elif t == "LI":
            body = [c for c in n["children"] if c["type"] != "L"]
            label = "".join(text_of(c) for c in body if c["type"] == "Lbl").strip()
            # The label decides bullet versus number; without one the list type is unknown and not scored.
            ordered = bool(re.match(r"^[(\[]?(\d+|[A-Za-z]|[ivxlcdm]+)[.)\]]", label, re.I)) if label else None
            doc.add_block(Block("list", split_words("".join(own_text(c) for c in body if c["type"] != "Lbl")), list_depth,
                                ordered))
            # Nested lists may sit directly under LI or inside its LBody.
            for c in nested_lists(n):
                walk(c, list_depth + 1)
        elif t == "L":
            for c in n["children"]:
                walk(c, list_depth)
        elif t == "Table":
            table = Table()
            for row in _descendants(n, "TR"):
                cells = [Cell(norm_text(text_of(c))) for c in row["children"] if c["type"] in ("TD", "TH")]
                if cells:
                    table.rows.append(cells)
            doc.add_table(table)
        elif t == "Figure":
            doc.images += 1
        elif t == "#text":
            doc.add_block(Block("paragraph", split_words(n["text"])))
        else:
            for c in n["children"]:
                walk(c, list_depth if t != "L" else list_depth)

    for c in root["children"]:
        walk(c, 0)
    return doc


def _descendants(n, kind):
    for c in n["children"]:
        if c["type"] == kind:
            yield c
        elif c["type"] not in ("Table",):
            yield from _descendants(c, kind)
