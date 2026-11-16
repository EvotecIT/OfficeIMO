"""Independent DOCX, XLSX and PPTX readers built on the package XML, without OfficeIMO or Office libraries."""
from __future__ import annotations

import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET

from .model import Block, Cell, Document, SheetCell, Table, Word, join_segments, norm_text, split_words
from .sheets import sheet_regions

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
W = "{%s}" % NS["w"]
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
S = "{%s}" % NS["s"]
R = "{%s}" % NS["r"]
MC = "{%s}" % NS["mc"]


def _xml(z: zipfile.ZipFile, name: str):
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    folder, name = posixpath.split(part)
    root = _xml(z, posixpath.join(folder, "_rels", name + ".rels"))
    out: dict[str, tuple[str, str]] = {}
    if root is None:
        return out
    for rel in root:
        target = rel.get("Target", "")
        if rel.get("TargetMode") != "External":
            target = posixpath.normpath(posixpath.join(folder, target)) if not target.startswith("/") else target[1:]
        out[rel.get("Id")] = (target, rel.get("TargetMode", ""))
    return out


def _on(el, attr=W + "val") -> bool:
    if el is None:
        return False
    return el.get(attr, "true").lower() not in ("0", "false", "off", "none")


def _external(url: str) -> bool:
    return url.lower().startswith(("http:", "https:", "mailto:", "ftp:"))


# ----------------------------------------------------------------------------------------------- DOCX

class _WordStyles:
    def __init__(self, z: zipfile.ZipFile):
        self.styles: dict[str, ET.Element] = {}
        self.default_size: float | None = None
        root = _xml(z, "word/styles.xml")
        if root is not None:
            sz = root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}sz")
            if sz is not None:
                self.default_size = int(sz.get(W + "val", "20")) / 2
            for style in root.findall(f"{W}style"):
                self.styles[style.get(W + "styleId")] = style
        self.formats: dict[tuple[str, int], str] = {}
        num = _xml(z, "word/numbering.xml")
        if num is not None:
            abstract = {}
            for an in num.findall(f"{W}abstractNum"):
                levels = {}
                for lvl in an.findall(f"{W}lvl"):
                    fmt = lvl.find(f"{W}numFmt")
                    levels[int(lvl.get(W + "ilvl", "0"))] = fmt.get(W + "val") if fmt is not None else "decimal"
                abstract[an.get(W + "abstractNumId")] = levels
            for n in num.findall(f"{W}num"):
                ref = n.find(f"{W}abstractNumId")
                levels = abstract.get(ref.get(W + "val") if ref is not None else None, {})
                for ilvl, fmt in levels.items():
                    self.formats[(n.get(W + "numId"), ilvl)] = fmt

    def chain(self, style_id: str | None):
        seen = set()
        while style_id and style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            style = self.styles[style_id]
            yield style
            based = style.find(f"{W}basedOn")
            style_id = based.get(W + "val") if based is not None else None

    def heading_level(self, style_id: str | None) -> int | None:
        for style in self.chain(style_id):
            name = (style.find(f"{W}name").get(W + "val") if style.find(f"{W}name") is not None else "").lower()
            if name == "title":
                return 0
            m = re.fullmatch(r"heading\s*([1-9])", name)
            if m:
                return int(m.group(1))
            lvl = style.find(f"{W}pPr/{W}outlineLvl")
            if lvl is not None and int(lvl.get(W + "val", "9")) < 9:
                return int(lvl.get(W + "val")) + 1
        return None

    def numbering(self, style_id: str | None):
        for style in self.chain(style_id):
            numpr = style.find(f"{W}pPr/{W}numPr")
            if numpr is not None:
                return numpr
        return None

    def run_prop(self, tag: str, style_ids: list[str | None]):
        for sid in style_ids:
            for style in self.chain(sid):
                el = style.find(f"{W}rPr/{W}{tag}")
                if el is not None:
                    return el
        return None


class _DocxReader:
    def __init__(self, z: zipfile.ZipFile):
        self.z = z
        self.st = _WordStyles(z)
        self.doc = Document()
        self.rels: dict[str, tuple[str, str]] = {}

    def read(self) -> Document:
        main = "word/document.xml"
        root = _xml(self.z, main)
        self.rels = _rels(self.z, main)
        body = root.find(f"{W}body")
        for target, _ in self.rels.values():
            if re.search(r"(header|footer)\d*\.xml$", target):
                part = _xml(self.z, target)
                if part is not None:
                    saved = self.rels
                    self.rels = _rels(self.z, target)
                    self._container(part, optional=True, count_images=False)
                    self.rels = saved
        self._container(body)
        for part, tag in (("word/footnotes.xml", "footnote"), ("word/endnotes.xml", "endnote")):
            notes = _xml(self.z, part)
            if notes is None:
                continue
            for note in notes.findall(f"{W}{tag}"):
                if note.get(W + "type") in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                text = norm_text(" ".join(self._para_text(p) for p in note.iter(f"{W}p")))
                if text:
                    self.doc.footnotes.append(text)
                    self.doc.footnotes_native = True
        for instr in root.iter(f"{W}instrText"):
            m = re.search(r'HYPERLINK\s+"([^"]+)"', instr.text or "")
            if m and _external(m.group(1)):
                self.doc.links.add(m.group(1))
        return self.doc

    def _container(self, el, optional=False, count_images=True):
        for child in el:
            tag = child.tag
            if tag == W + "p":
                self._paragraph(child, optional, count_images)
            elif tag == W + "tbl":
                if optional:
                    for p in child.iter(W + "p"):
                        self._paragraph(p, True, False)
                else:
                    self.doc.add_table(self._table(child))
            elif tag in (W + "sdt", W + "customXml"):
                content = child.find(f"{W}sdtContent")
                self._container(content if content is not None else child, optional, count_images)

    def _table(self, tbl) -> Table:
        table = Table()
        open_cells: dict[int, Cell] = {}
        for tr in tbl.findall(f"{W}tr"):
            row: list[Cell] = []
            col = 0
            before = tr.find(f"{W}trPr/{W}gridBefore")
            if before is not None:
                col += int(before.get(W + "val", "0"))
            for tc in tr.findall(f"{W}tc"):
                pr = tc.find(f"{W}tcPr")
                span = 1
                vmerge = None
                if pr is not None:
                    gs = pr.find(f"{W}gridSpan")
                    span = int(gs.get(W + "val", "1")) if gs is not None else 1
                    vm = pr.find(f"{W}vMerge")
                    if vm is not None:
                        vmerge = vm.get(W + "val", "continue")
                if vmerge == "continue" and col in open_cells:
                    open_cells[col].row_span += 1
                else:
                    text = " ".join(self._para_text(p) for p in tc.iter(f"{W}p"))
                    cell = Cell(norm_text(text), 1, span)
                    row.append(cell)
                    for c in range(col, col + span):
                        open_cells.pop(c, None)
                    if vmerge == "restart":
                        open_cells[col] = cell
                for image in tc.iter():
                    if image.tag in (A + "blip", "{%s}imagedata" % NS["v"]):
                        self.doc.images += 1
                for link in tc.iter(f"{W}hyperlink"):
                    self._link(link)
                col += span
            if row:
                table.rows.append(row)
        return table

    def _link(self, el):
        rid = el.get(R + "id")
        if rid and rid in self.rels:
            target = self.rels[rid][0]
            if _external(target):
                self.doc.links.add(target)

    def _para_text(self, p) -> str:
        return " ".join(w.text for w in self._runs(p))

    def _runs(self, p, boxes=None, images=None) -> list[Word]:
        segments: list[tuple[str, bool, bool, float | None]] = []
        ppr = p.find(f"{W}pPr")
        pstyle = ppr.find(f"{W}pStyle").get(W + "val") if ppr is not None and ppr.find(f"{W}pStyle") is not None else None

        def walk(el):
            for child in el:
                tag = child.tag
                if tag == W + "r":
                    run(child)
                elif tag in (W + "hyperlink", W + "ins", W + "smartTag", W + "fldSimple", W + "customXml", W + "bdo", W + "dir"):
                    if tag == W + "hyperlink":
                        self._link(child)
                    walk(child)
                elif tag == W + "sdt":
                    content = child.find(f"{W}sdtContent")
                    if content is not None:
                        walk(content)
                elif tag == MC + "AlternateContent":
                    choice = child.find(f"{MC}Choice")
                    if choice is not None:
                        walk(choice)

        def run(r):
            rpr = r.find(f"{W}rPr")
            rstyle = rpr.find(f"{W}rStyle").get(W + "val") if rpr is not None and rpr.find(f"{W}rStyle") is not None else None
            if rpr is not None and _on(rpr.find(f"{W}vanish")):
                return

            def prop(tag):
                el = rpr.find(f"{W}{tag}") if rpr is not None else None
                return el if el is not None else self.st.run_prop(tag, [rstyle, pstyle])

            bold, italic = _on(prop("b")), _on(prop("i"))
            sz = prop("sz")
            size = int(sz.get(W + "val")) / 2 if sz is not None and sz.get(W + "val", "").isdigit() else self.st.default_size
            parts: list[str] = []
            for child in r:
                tag = child.tag
                if tag == W + "t":
                    parts.append(child.text or "")
                elif tag in (W + "tab", W + "br", W + "cr"):
                    parts.append(" ")
                elif tag == W + "noBreakHyphen":
                    parts.append("-")
                elif tag in (W + "drawing", W + "pict", W + "object", MC + "AlternateContent"):
                    _scan_graphic(child, images, boxes)
            segments.append(("".join(parts), bold, italic, size))

        walk(p)
        return join_segments(segments)

    def _paragraph(self, p, optional=False, count_images=True):
        boxes: list = []
        images: list = []
        words = self._runs(p, boxes, images)
        if count_images:
            self.doc.images += len(images)
        ppr = p.find(f"{W}pPr")
        pstyle = None
        numpr = None
        level = None
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            pstyle = ps.get(W + "val") if ps is not None else None
            numpr = ppr.find(f"{W}numPr")
            ol = ppr.find(f"{W}outlineLvl")
            if ol is not None and int(ol.get(W + "val", "9")) < 9:
                level = int(ol.get(W + "val")) + 1
        if level is None:
            level = self.st.heading_level(pstyle)
        if numpr is None:
            numpr = self.st.numbering(pstyle)
        block = Block("paragraph", words, optional=optional)
        if level is not None:
            block.kind, block.level = "heading", level
        elif numpr is not None:
            num = numpr.find(f"{W}numId")
            ilvl = numpr.find(f"{W}ilvl")
            num_id = num.get(W + "val") if num is not None else None
            lvl = int(ilvl.get(W + "val", "0")) if ilvl is not None else 0
            if ilvl is None:
                # Built-in "List Bullet 2"-style paragraphs nest through separate numbering rather than ilvl.
                for style in self.st.chain(pstyle):
                    name = style.find(f"{W}name")
                    m = re.fullmatch(r"list (?:bullet|number|continue)?\s*([2-9])", (name.get(W + "val") if name is not None else "").lower())
                    if m:
                        lvl = int(m.group(1)) - 1
                        break
            fmt = self.st.formats.get((num_id, int(ilvl.get(W + "val", "0")) if ilvl is not None else 0), "decimal")
            if num_id not in (None, "0") and fmt != "none":
                block.kind, block.level, block.ordered = "list", lvl, fmt != "bullet"
        self.doc.add_block(block)
        for box in boxes:
            self._container(box, optional, count_images)


def _scan_graphic(el, images, boxes) -> None:
    """Counts pictures and collects text-box bodies, ignoring markup-compatibility fallbacks that duplicate them."""
    for child in el:
        tag = child.tag
        if tag == MC + "Fallback":
            continue
        if tag in (A + "blip", "{%s}imagedata" % NS["v"]):
            if images is not None:
                images.append(1)
        elif tag == W + "txbxContent":
            if boxes is not None:
                boxes.append(child)
            continue
        _scan_graphic(child, images, boxes)


def read_docx(path: str) -> Document:
    with zipfile.ZipFile(path) as z:
        return _DocxReader(z).read()


# ----------------------------------------------------------------------------------------------- XLSX

_BUILTIN_FORMATS = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
                    12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
                    19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)",
                    38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
                    46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@"}


def _format_info(code: str) -> tuple[bool, int | None, bool]:
    """(is_date, decimals, percent) for the first section of a number format."""
    section = code.split(";")[0]
    stripped = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.', "", section)
    is_date = bool(re.search(r"[dmyhs]", stripped, re.I)) and not re.search(r"[0#?]", stripped)
    if code.lower() == "general" or code == "@":
        return False, None, False
    m = re.search(r"\.([0#?]+)", stripped)
    decimals = len(m.group(1)) if m else (0 if re.search(r"[0#?]", stripped) else None)
    return is_date, decimals, "%" in stripped


def _col_index(ref: str) -> tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


def read_xlsx(path: str) -> Document:
    doc = Document()
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        sst = _xml(z, "xl/sharedStrings.xml")
        if sst is not None:
            for si in sst.findall(f"{S}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{S}t") if not _in_phonetic(si, t)))
        formats: list[tuple[bool, int | None, bool]] = []
        styles = _xml(z, "xl/styles.xml")
        if styles is not None:
            custom = {int(n.get("numFmtId")): n.get("formatCode", "") for n in styles.iter(f"{S}numFmt")}
            xfs = styles.find(f"{S}cellXfs")
            for xf in (xfs if xfs is not None else []):
                fid = int(xf.get("numFmtId", "0"))
                code = custom.get(fid, _BUILTIN_FORMATS.get(fid, "General"))
                info = _format_info(code)
                if fid in range(14, 23) or fid in (45, 46, 47):
                    info = (True, info[1], False)
                formats.append(info)
        wb = _xml(z, "xl/workbook.xml")
        rels = _rels(z, "xl/workbook.xml")
        row_offset = 0
        for sheet in wb.iter(f"{S}sheet"):
            if sheet.get("state") in ("hidden", "veryHidden"):
                continue
            part = rels.get(sheet.get(R + "id"), ("", ""))[0]
            root = _xml(z, part)
            if root is None:
                continue
            srels = _rels(z, part)
            cells: dict[tuple[int, int], SheetCell] = {}
            for c in root.iter(f"{S}c"):
                r, col = _col_index(c.get("r"))
                t = c.get("t", "n")
                v = c.find(f"{S}v")
                raw = v.text if v is not None else None
                style = int(c.get("s", "0"))
                is_date, decimals, percent = formats[style] if style < len(formats) else (False, None, False)
                if t == "s" and raw is not None:
                    cell = SheetCell(r, col, shared[int(raw)], "text")
                elif t == "inlineStr":
                    cell = SheetCell(r, col, "".join(x.text or "" for x in c.iter(f"{S}t")), "text")
                elif t == "str":
                    cell = SheetCell(r, col, raw or "", "text")
                elif t == "b":
                    cell = SheetCell(r, col, "TRUE" if raw == "1" else "FALSE", "bool")
                elif t == "e":
                    cell = SheetCell(r, col, raw or "", "error")
                elif t == "d":
                    cell = SheetCell(r, col, raw or "", "date")
                elif raw is not None:
                    value = float(raw)
                    if is_date:
                        cell = SheetCell(r, col, _serial_to_iso(value), "date", value)
                    else:
                        # The PDF shows the rounded value, so the ground-truth text is rounded the same way.
                        shown = round(value, decimals + (2 if percent else 0)) if decimals is not None else value
                        cell = SheetCell(r, col, format(shown, ".15g"), "number", value, decimals, percent)
                else:
                    continue
                if norm_text(cell.text):
                    cells[(r, col)] = cell
            merges = []
            for m in root.iter(f"{S}mergeCell"):
                a, _, b = m.get("ref").partition(":")
                (r1, c1), (r2, c2) = _col_index(a), _col_index(b or a)
                merges.append((r1, c1, r2, c2))
            for cell in cells.values():
                doc.cells.append(SheetCell(cell.row + row_offset, cell.col, cell.text, cell.kind, cell.value, cell.decimals, cell.percent))
            sheet_regions(doc, cells, merges)
            if cells:
                row_offset += max(r for r, _ in cells) + 2
            for h in root.iter(f"{S}hyperlink"):
                rid = h.get(R + "id")
                if rid in srels and _external(srels[rid][0]):
                    doc.links.add(srels[rid][0])
            for target, _ in srels.values():
                if "drawings/" in target:
                    drawing = _xml(z, target)
                    if drawing is not None:
                        doc.images += sum(1 for _ in drawing.iter("{%s}pic" % NS["xdr"]))
    return doc


def _serial_to_iso(serial: float) -> str:
    import datetime as dt

    try:
        return (dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))).isoformat()
    except OverflowError:
        return str(serial)


def _in_phonetic(si, t) -> bool:
    for rph in si.findall(f"{S}rPh"):
        if t in list(rph.iter()):
            return True
    return False


# ----------------------------------------------------------------------------------------------- PPTX

def read_pptx(path: str) -> Document:
    doc = Document()
    with zipfile.ZipFile(path) as z:
        pres = _xml(z, "ppt/presentation.xml")
        rels = _rels(z, "ppt/presentation.xml")
        slides = []
        for sid in pres.iter(f"{P}sldId"):
            slides.append(rels[sid.get(R + "id")][0])
        doc.pages = 0
        sizes = _master_sizes(z)
        for part in slides:
            root = _xml(z, part)
            if root is None or root.get("show") == "0":
                continue
            doc.pages += 1
            srels = _rels(z, part)
            tree = root.find(f"{P}cSld/{P}spTree")
            _pptx_tree(doc, tree, srels, sizes)
    return doc


def _master_sizes(z: zipfile.ZipFile) -> dict:
    """Default run sizes from the first slide master's text styles (layout overrides are not resolved)."""
    master = _xml(z, "ppt/slideMasters/slideMaster1.xml")
    out = {}
    if master is None:
        return out
    for key, tag in (("title", "titleStyle"), ("body", "bodyStyle"), ("other", "otherStyle")):
        style = master.find(f"{P}txStyles/{P}{tag}")
        if style is None:
            continue
        for lvl in range(9):
            d = style.find(f"{A}lvl{lvl + 1}pPr/{A}defRPr")
            if d is not None and d.get("sz"):
                out[(key, lvl)] = int(d.get("sz")) / 100
    return out


def _pptx_tree(doc: Document, tree, rels, sizes) -> None:
    for el in tree:
        tag = el.tag
        if tag == P + "grpSp":
            _pptx_tree(doc, el, rels, sizes)
        elif tag == P + "pic":
            doc.images += 1
        elif tag == P + "sp":
            ph = el.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
            ph_type = ph.get("type", "body") if ph is not None else None
            body = el.find(f"{P}txBody")
            if body is None:
                continue
            for para in body.findall(f"{A}p"):
                segments: list[tuple[str, bool, bool, float | None]] = []
                ppr0 = para.find(f"{A}pPr")
                plvl = int(ppr0.get("lvl", "0")) if ppr0 is not None else 0
                kind = "title" if ph_type in ("title", "ctrTitle") else "body" if ph_type is not None else "other"
                default_size = sizes.get((kind, plvl))
                for r in para:
                    if r.tag in (A + "r", A + "fld"):
                        rpr = r.find(f"{A}rPr")
                        bold = rpr is not None and rpr.get("b") in ("1", "true")
                        italic = rpr is not None and rpr.get("i") in ("1", "true")
                        size = int(rpr.get("sz")) / 100 if rpr is not None and rpr.get("sz") else default_size
                        if rpr is not None:
                            link = rpr.find(f"{A}hlinkClick")
                            if link is not None and link.get(R + "id") in rels and _external(rels[link.get(R + "id")][0]):
                                doc.links.add(rels[link.get(R + "id")][0])
                        t = r.find(f"{A}t")
                        segments.append((t.text if t is not None and t.text else "", bold, italic, size))
                    elif r.tag == A + "br":
                        segments.append((" ", False, False, None))
                words = join_segments(segments)
                ppr = para.find(f"{A}pPr")
                lvl = int(ppr.get("lvl", "0")) if ppr is not None else 0
                block = Block("paragraph", words)
                if ph_type in ("title", "ctrTitle"):
                    block.kind, block.level = "heading", 1
                else:
                    explicit_none = ppr is not None and ppr.find(f"{A}buNone") is not None
                    bullet = ppr is not None and (ppr.find(f"{A}buChar") is not None or ppr.find(f"{A}buAutoNum") is not None)
                    if bullet or (ph_type in ("body", "obj") and not explicit_none):
                        block.kind, block.level = "list", lvl
                        block.ordered = ppr is not None and ppr.find(f"{A}buAutoNum") is not None
                doc.add_block(block)
        elif tag == P + "graphicFrame":
            tbl = el.find(f".//{A}tbl")
            if tbl is None:
                continue
            table = Table()
            for tr in tbl.findall(f"{A}tr"):
                row = []
                for tc in tr.findall(f"{A}tc"):
                    if tc.get("hMerge") in ("1", "true") or tc.get("vMerge") in ("1", "true"):
                        continue
                    text = " ".join("".join(t.text or "" for t in p.iter(f"{A}t")) for p in tc.iter(f"{A}p"))
                    row.append(Cell(norm_text(text), int(tc.get("rowSpan", "1")), int(tc.get("gridSpan", "1"))))
                if row:
                    table.rows.append(row)
            doc.add_table(table)
