#!/usr/bin/env python3
"""Render OfficeIMO PowerPoint layout mockups from powerpoint-layouts.json and the theme.

Usage:
  python3 Build/Visual/powerpoint-mockups.py [--out DIR] [--themes Modern,Boardroom] [--write-abs] [--no-render]

Geometry comes only from powerpoint-layouts.json; colours, fonts and the base type
scale come from Build/Visual/themes.json when it exists, otherwise from the
provisional values in OfficeIMO.Core/Themes/OfficeTheme.BuiltIns.cs. Re-run after
either file changes; nothing in this script hard-codes a colour or a box.

--write-abs rewrites the "abs" boxes next to every "grid" reference in the JSON.
Without it the script fails when a stored "abs" box disagrees with the grid.
"""
import argparse, html, json, math, os, re, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LAYOUTS = ROOT / "Build/Visual/powerpoint-layouts.json"
THEMES_JSON = ROOT / "Build/Visual/themes.json"
BUILTINS = ROOT / "OfficeIMO.Core/Themes/OfficeTheme.BuiltIns.cs"
DEFAULT_OUT = Path("/Volumes/Samsung/Development/Scratch/officeimo-visual-design/powerpoint")
EDGE = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
FONT_DIR = Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts")

# ---------------------------------------------------------------- geometry

def pt(v):
    if isinstance(v, (int, float)): return float(v)
    m = re.fullmatch(r"\s*(-?[\d.]+)\s*(pt|in)?\s*", str(v))
    if not m: raise ValueError(v)
    n = float(m.group(1)); return n * 72 if m.group(2) == "in" else n

class Grid:
    def __init__(self, size):
        g = size["grid"]
        self.W, self.H = size["widthPt"], size["heightPt"]
        self.cw, self.g = pt(g["column"]), pt(g["gutter"])
        self.rh, self.rg = pt(g["row"]), pt(g["rowGutter"])
        m = g["margin"]
        val = lambda v: pt(v["default"]) if isinstance(v, dict) else pt(v)
        self.ml, self.mt, self.mr = val(m["left"]), val(m["top"]), val(m["right"])
        self.live_bottom = self.H - pt(m["bottom"])
        self.footer_y, self.footer_h = pt(size["footerBand"]["y"]), pt(size["footerBand"]["h"])
    def for_theme(self, t):
        """Side/top margins come from the theme's spacing tokens; columns and rows re-derive."""
        import copy
        g = copy.copy(self); sp = t.get("spacing", {})
        g.ml = sp.get("slideMarginLeft", self.ml); g.mr = sp.get("slideMarginRight", self.mr); g.mt = sp.get("slideMarginTop", self.mt)
        g.cw = (self.W - g.ml - g.mr - 11 * self.g) / 12
        g.rh = (self.live_bottom - g.mt - 11 * self.rg) / 12
        return g
    def x(self, c): return self.ml + (c - 1) * (self.cw + self.g)
    def w(self, s): return s * self.cw + (s - 1) * self.g
    def y(self, r): return self.mt + (r - 1) * (self.rh + self.rg)
    def h(self, n): return n * self.rh + (n - 1) * self.rg
    def box(self, col, span, row, rows): return (self.x(col), self.y(row), self.w(span), self.h(rows))
    def gbox(self, gr): return self.box(gr["col"], gr["span"], gr["row"], gr["rows"])

def fmt(n): return ("%g" % round(n, 2)) + "pt"

def sync_abs(spec_text, grid, write):
    pat = re.compile(r'"grid": \{ "col": (\d+), "span": (\d+), "row": (\d+), "rows": (\d+) \}(, "abs": \{[^}]*\})?')
    bad = []
    def rep(m):
        c, s, r, n = map(int, m.groups()[:4])
        x, y, w, h = grid.box(c, s, r, n)
        a = '"abs": { "x": "%s", "y": "%s", "w": "%s", "h": "%s" }' % (fmt(x), fmt(y), fmt(w), fmt(h))
        if m.group(5) and m.group(5)[2:] != a: bad.append(m.group(0))
        if not m.group(5): bad.append(m.group(0))
        return m.group(0)[: m.end(4) - m.start(0) + 2] + ", " + a
    new = pat.sub(rep, spec_text)
    if write:
        LAYOUTS.write_text(new); print("abs boxes written:", len(pat.findall(spec_text)))
    elif bad:
        sys.exit("abs boxes out of date (%d); run with --write-abs" % len(bad))

# ---------------------------------------------------------------- theme

def hexc(s): s = s.strip().lstrip("#"); return "#" + s.upper()

def parse_builtins():
    src = BUILTINS.read_text()
    themes = {}
    body = src[src.index("private static OfficeTheme Create"):]
    def pick(expr, dark):
        m = re.fullmatch(r'C\(dark \? "(\w+)" : "(\w+)"\)', expr.strip())
        if m: return hexc(m.group(1) if dark else m.group(2))
        return hexc(re.fullmatch(r'C\("(\w+)"\)', expr.strip()).group(1))
    pal_args = re.search(r"new OfficeThemePalette\((.*?)\);", body, re.S).group(1)
    pal_exprs = re.findall(r'C\((?:dark \? "\w+" : "\w+"|"\w+"|background|text|muted|accent)\)', pal_args)
    scale = [float(v) for v in re.findall(r"P\(([\d.]+|size)\)", re.search(r"new OfficeThemeTypeScale\((.*?)\);", body).group(1)) if v != "size"]
    tab = re.search(r"new OfficeThemeTableStyle\((.*?)\);", body, re.S).group(1)
    tab_c = re.findall(r'C\(dark \? "\w+" : "\w+"\)', tab)
    cell = float(re.findall(r"P\(([\d.]+)\)", tab)[-1])
    for m in re.finditer(r'public static OfficeTheme (\w+) \{ get; \} = Create\("(\w+)", "(\w+)", "(\w+)", "(\w+)", "(\w+)", "([^"]+)", "([^"]+)", ([\d.]+)\);', src):
        name, bg, tx, mu, ac, hf, bf, size = m.group(2), *m.groups()[2:]
        dark = name == "Dark"
        lit = {"background": hexc(bg), "text": hexc(tx), "muted": hexc(mu), "accent": hexc(ac)}
        vals = []
        for e in pal_exprs:
            inner = re.fullmatch(r"C\((\w+)\)", e)
            vals.append(lit[inner.group(1)] if inner else pick(e, dark))
        keys = ["background", "text", "textMuted", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "success", "warning", "danger"]
        t = {"name": name, "palette": dict(zip(keys, vals)), "fonts": {"heading": hf, "body": bf},
             "typeScale": dict(zip(["title", "heading1", "heading2", "heading3", "heading4", "heading5", "heading6"], scale[:7])),
             "tableStyle": {"headerFill": "accent1", "headerText": "#FFFFFF", "bandFill": pick(tab_c[0], dark), "borderColor": pick(tab_c[1], dark), "cellPadding": cell},
             "source": str(BUILTINS.relative_to(ROOT))}
        t["typeScale"]["body"] = float(size); t["typeScale"]["caption"] = scale[7] if len(scale) > 7 else 9.0
        themes[name] = t
    return themes

def find_ci(d, *names):
    if not isinstance(d, dict): return None
    low = {k.lower(): v for k, v in d.items()}
    for n in names:
        if n.lower() in low: return low[n.lower()]
    return None

def size_of(v):
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, str): return pt(v)
    if isinstance(v, dict): return size_of(find_ci(v, "value", "size", "points", "pt"))
    return None

def color_of(v):
    if isinstance(v, str) and re.fullmatch(r"#?[0-9A-Fa-f]{6}", v.strip()): return hexc(v)
    if isinstance(v, str): return v  # role reference such as "accent1"
    if isinstance(v, dict): return color_of(find_ci(v, "hex", "value", "color"))
    return None

def load_themes():
    base = parse_builtins()
    if not THEMES_JSON.exists():
        return base, "provisional values from " + str(BUILTINS.relative_to(ROOT))
    data = json.loads(THEMES_JSON.read_text())
    found = {}
    def walk(o):
        if isinstance(o, dict):
            nm = o.get("name") or o.get("Name")
            if isinstance(nm, str) and nm in base and find_ci(o, "palette", "colors"): found[nm] = o
            for k, v in o.items():
                if k in base and isinstance(v, dict) and find_ci(v, "palette", "colors"): found[k] = v
                walk(v)
        elif isinstance(o, list):
            for v in o: walk(v)
    walk(data)
    for name, o in found.items():
        t = base[name]; t["source"] = "Build/Visual/themes.json"
        pal = find_ci(o, "palette", "colors") or {}
        for k, alts in {"background": ["background", "bg"], "text": ["text", "foreground"], "textMuted": ["textMuted", "muted", "mutedText"],
                        "accent1": ["accent1"], "accent2": ["accent2"], "accent3": ["accent3"], "accent4": ["accent4"], "accent5": ["accent5"], "accent6": ["accent6"],
                        "success": ["success"], "warning": ["warning"], "danger": ["danger"]}.items():
            c = color_of(find_ci(pal, *alts))
            if c and c.startswith("#"): t["palette"][k] = c
        fonts = find_ci(o, "fonts") or {}
        for k in ("heading", "body"):
            f = find_ci(fonts, k)
            if isinstance(f, dict): f = find_ci(f, "name", "family", "latin")
            if isinstance(f, str): t["fonts"][k] = f
        ts = find_ci(o, "typeScale", "typescale", "sizes") or {}
        for k in list(t["typeScale"]):
            v = size_of(find_ci(ts, k))
            if v: t["typeScale"][k] = v
        sp = find_ci(o, "spacing") or {}
        t["spacing"] = {k: size_of(v) for k, v in sp.items() if size_of(v) is not None}
        ext = find_ci(o, "extensions") or {}
        if color_of(find_ci(ext, "surface")): t["tableStyle"]["surface"] = color_of(find_ci(ext, "surface"))
        if color_of(find_ci(ext, "rule")): t["tableStyle"]["rule"] = color_of(find_ci(ext, "rule"))
        hr = (find_ci(ext, "table") or {}).get("headerRule")
        t["tableStyle"]["headerRule"] = hr
        tb = find_ci(o, "tableStyle", "table") or {}
        for k in ("bandFill", "borderColor", "headerText", "headerFill"):
            c = color_of(find_ci(tb, k))
            if c: t["tableStyle"][k] = c
        v = size_of(find_ci(tb, "cellPadding"))
        if v: t["tableStyle"]["cellPadding"] = v
    return base, "Build/Visual/themes.json (%s), others provisional" % ", ".join(sorted(found)) if found else "themes.json present but unreadable; provisional values"

def lum(h):
    h = h.lstrip("#"); r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)
def contrast(a, b):
    la, lb = sorted([lum(a), lum(b)], reverse=True); return (la + 0.05) / (lb + 0.05)
def rgba(h, a):
    h = h.lstrip("#"); return "rgba(%d,%d,%d,%.3f)" % (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), a)

def round_size(v): return float(round(v / 2) * 2) if v >= 18 else float(round(v))

class Theme:
    def __init__(self, t, spec):
        self.t = t; self.name = t["name"]; p = t["palette"]; self.p = p
        tb = t["tableStyle"]
        self.surface = tb.get("surface", tb["bandFill"]); self.hairline = tb.get("rule", tb["borderColor"])
        self.header_fill = self.role(tb["headerFill"]); self.header_text = self.role(tb["headerText"])
        hr = tb.get("headerRule")
        self.header_rule = (self.role(hr["color"]), hr.get("width", 1)) if hr else None
        self.cell_pad = tb["cellPadding"]
        self.ink, self.on_ink = (p["text"], p["background"]) if lum(p["text"]) < lum(p["background"]) else (p["background"], p["text"])
        self.on_accent = "#FFFFFF" if contrast(p["accent1"], "#FFFFFF") >= 4.5 else self.ink
        self.styles = {}
        for k, s in spec["textStyles"].items():
            if k.startswith("$"): continue
            size = round_size(t["typeScale"][s["size"]["token"].split(".")[1]] * s["size"]["factor"])
            self.styles[k] = dict(s, pt=size, minPt=pt(s["min"]))
        self.heading_font, self.body_font = t["fonts"]["heading"], t["fonts"]["body"]
        # Average advance width relative to Aptos; the engine measures real glyph advances instead.
        wf = {"Georgia": 1.14, "Cambria": 1.05, "Arial": 1.05, "Aptos": 1.0, "Aptos Display": 0.97, "Calibri": 0.93}
        self.hfac, self.bfac = wf.get(self.heading_font, 1.05), wf.get(self.body_font, 1.05)
    def role(self, v):
        v = str(v)
        if v.startswith("#"): return v
        return self.p[v[0].lower() + v[1:]]
    def color(self, role):
        role = role.split(" ")[0]
        if role.startswith("palette."): return self.p[role[8:]]
        return {"surface": self.surface, "hairline": self.hairline, "ink": self.ink, "onInk": self.on_ink, "onAccent": self.on_accent,
                "onAccentMuted": rgba(self.on_accent, .72), "onInkMuted": rgba(self.on_ink, .64), "tableStyle.headerText": self.header_text}[role]

# ---------------------------------------------------------------- measurement (approximation of font metrics)

def char_w(ch):
    if ch == " ": return 0.26
    if ch in "il.,:;'|!": return 0.26
    if ch in "frtj()-": return 0.36
    if ch in "mwMW": return 0.82
    if ch.isdigit(): return 0.56
    if ch.isupper(): return 0.64
    return 0.51

def text_w(s, size, tracking=0.0): return sum(char_w(c) + tracking for c in s) * size

def wrap(s, size, width, tracking=0.0):
    lines, cur = [], ""
    for word in s.split(" "):
        t = word if not cur else cur + " " + word
        if text_w(t, size, tracking) <= width or not cur: cur = t
        else: lines.append(cur); cur = word
    if cur: lines.append(cur)
    return lines

# ---------------------------------------------------------------- html primitives

E = html.escape

class Slide:
    def __init__(self, th, grid, layout, bg):
        self.th, self.g, self.layout, self.bg, self.parts = th, grid, layout, bg, []
    def add(self, s): self.parts.append(s)
    def box(self, x, y, w, h, inner, css=""):
        self.add('<div class="b" style="left:%gpt;top:%gpt;width:%gpt;height:%gpt;%s">%s</div>' % (x, y, w, h, css, inner))
    def style_css(self, key, color=None, size=None, extra=""):
        s = self.th.styles[key]
        fam = self.th.heading_font if s["font"] == "fonts.heading" else self.th.body_font
        col = color or self.th.color(s["color"]) if not (s["color"].startswith("onInk @") ) else rgba(self.th.on_ink, .09)
        if color: col = color
        css = "font-family:'%s',Aptos,Arial,sans-serif;font-size:%gpt;font-weight:%d;line-height:%s;color:%s;" % (fam, size or s["pt"], s["weight"], s["lineHeight"], col)
        if "tracking" in s: css += "letter-spacing:%s;" % s["tracking"]
        if s.get("case") == "upper": css += "text-transform:uppercase;"
        if s.get("numerals") == "tabular": css += "font-variant-numeric:tabular-nums;"
        return css + extra
    def text(self, x, y, w, h, key, content, anchor="top", align="left", color=None, size=None, extra="", raw=False):
        jc = {"top": "flex-start", "bottom": "flex-end", "middle": "center"}[anchor]
        body = content if raw else E(content)
        self.box(x, y, w, h, '<div style="%s">%s</div>' % (self.style_css(key, color, size, "text-align:%s;" % align + extra), body),
                 "display:flex;flex-direction:column;justify-content:%s;" % jc)
    def html(self, title):
        th = self.th
        faces = font_faces()
        return """<!doctype html><html><head><meta charset="utf-8"><title>%s</title><style>%s
html,body{margin:0;padding:0;background:#777}html{zoom:1.25}
.s{position:relative;width:%gpt;height:%gpt;overflow:hidden;background:%s}
.b{position:absolute;box-sizing:border-box}
.b>div{-webkit-font-smoothing:antialiased}
ul{margin:0;padding:0;list-style:none}
table{border-collapse:collapse;width:100%%}
</style></head><body><div class="s">%s</div></body></html>""" % (E(title), faces, self.g.W, self.g.H, self.bg, "\n".join(self.parts))

_faces = None
def font_faces():
    global _faces
    if _faces is not None: return _faces
    css = []
    if FONT_DIR.exists():
        for fam in ("Aptos", "Aptos Display"):
            for w, f in ((400, "Aptos.ttf"), (500, "Aptos.ttf"), (600, "Aptos-SemiBold.ttf"), (700, "Aptos-Bold.ttf"), (800, "Aptos-ExtraBold.ttf"), (300, "Aptos-Light.ttf")):
                if (FONT_DIR / f).exists():
                    css.append("@font-face{font-family:'%s';font-weight:%d;src:url('file://%s')}" % (fam, w, FONT_DIR / f))
        for fam, files in (("Calibri", ("Calibri.ttf", "Calibrib.ttf")), ("Cambria", ("Cambria.ttc", "Cambriab.ttf"))):
            for w, f in zip((400, 700), files):
                if (FONT_DIR / f).exists(): css.append("@font-face{font-family:'%s';font-weight:%d;src:url('file://%s')}" % (fam, w, FONT_DIR / f))
    _faces = "\n".join(css); return _faces

# ---------------------------------------------------------------- layout engine

class Deck:
    def __init__(self, spec, th, grid, title, footer=None):
        self.spec, self.th, self.g, self.title, self.footer = spec, th, grid, title, footer or title
        self.slides, self.section, self.section_no = [], None, 0

    # --- selection (spec.selection.rules, in order)
    def select(self, s, index, count):
        if s.get("layout"): return s["layout"]
        b = blocks(s)
        kinds = [k for k, _ in b]
        if index == 0 and not b: return "Title"
        if "location" in kinds: return "Coverage"
        if "kpi" in kinds and "card" in kinds: return "CaseStudy"
        if "kpi" in kinds: return "Kpi"
        if "card" in kinds:
            cards = dict(b)["card"]
            return "Logos" if all(c.get("image") and not c.get("text") and not c.get("items") for c in cards) else "Cards"
        if "step" in kinds: return "Process"
        if len(b) == 2 and all(k == "column" for k in kinds) and all(v.get("title") for _, v in b): return "Comparison"
        if len(b) == 1 and kinds[0] == "image" and b[0][1].get("fullBleed"): return "ImageFull"
        if not s.get("title") and len(b) == 1 and kinds[0] == "text" and len(b[0][1]) <= 240: return "Quote"
        if index == count - 1 and count >= 5 and (not b or (kinds == ["list"] and len(b[0][1]["items"]) <= 4)): return "Closing"
        if not b: return "Section"
        if len(b) == 2: return "TwoContent"
        return "TitleAndContent"

    def build(self, sources):
        out = []
        for i, s in enumerate(sources):
            lay = self.select(s, i, len(sources))
            out.extend(self.expand(dict(s, _layout=lay)))
        return out

    # --- overflow: returns one or more slide specs
    def expand(self, s):
        lay = s["_layout"]
        if lay == "TitleAndContent":
            kind, val = blocks(s)[0]
            if kind == "list": return self.split_list(s, val)
            if kind == "table": return self.split_table(s, val)
        return [s]

    def header_geom(self, s):
        """chrome.header: title flows from row 2, the subtitle follows the title, and the
        content zone starts on the first row at least 36pt below the last header line."""
        g, th = self.g, self.th
        if "_titlePt" not in s: self.title_fit(s)
        size = s["_titlePt"]
        title = s.get("title", ""); suffix = s.get("_suffix", "")
        lines = len(wrap(title, size * th.hfac, g.w(10)))
        sub_pt = th.styles["subtitle"]["pt"]
        if suffix:
            # The continued suffix rides on the title's last line at subtitle size; if it would add a
            # line it moves to the eyebrow line instead, so a continuation never changes the header.
            last = wrap(title, size * th.hfac, g.w(10))[-1]
            fits = (text_w(last, size * th.hfac) + text_w(suffix, sub_pt * th.bfac)) * 1.1 <= g.w(10)
            s["_suffixPlace"] = "title" if fits else "eyebrow"
        bottom = g.y(2) + lines * size * th.styles["title"]["lineHeight"]
        geo = {"titleY": g.y(2), "titleLines": lines, "titlePt": size}
        if s.get("subtitle"):
            sp = th.styles["subtitle"]["pt"]; geo["subtitleY"] = bottom + 10
            bottom = geo["subtitleY"] + len(wrap(s["subtitle"], sp * th.bfac, g.w(9))) * sp * th.styles["subtitle"]["lineHeight"]
        r0 = next(r for r in range(4, 9) if g.y(r) >= bottom + 36)
        geo["r0"], geo["n"] = r0, 13 - r0
        return geo

    def content_zone(self, s):
        geo = self.header_geom(s); return geo["r0"], geo["n"]

    def title_fit(self, s):
        """overflow.title steps; mutates s with resolved title size/lines."""
        st = self.th.styles["title"]; w = self.g.w(10)
        t = s.get("title", "")
        f = self.th.hfac
        if len(wrap(t, st["pt"] * f, w)) <= 2: s["_titlePt"] = st["pt"]; return
        for sep in (": ", " — ", " -- ", " - "):
            if sep in t and not s.get("subtitle"):
                a, b = t.split(sep, 1); s["title"], s["subtitle"] = a, b[0].upper() + b[1:]
                return self.title_fit(s)
        size = st["pt"]
        while size - 2 >= st["minPt"] and len(wrap(t, size * f, w)) > 2: size -= 2
        if len(wrap(t, size * f, w)) > 2:
            size = st["minPt"]; s["_longTitle"] = True; s.setdefault("_diag", []).append("title-long")
        s["_titlePt"] = size

    def list_height(self, items, size, width):
        body = self.th.styles["body"]; gap = 8 * size / body["pt"]; h = 0
        for it in items:
            lvl = it[1] if isinstance(it, tuple) else 1
            txt = it[0] if isinstance(it, tuple) else it
            sz = size if lvl == 1 else size * 0.86
            ind = 18 + (0 if lvl == 1 else 22)
            h += len(wrap(txt, sz * self.th.bfac, width - ind)) * sz * body["lineHeight"] + gap
        return h - gap

    def split_list(self, s, val):
        self.title_fit(s)
        r0, n = self.content_zone(s); boxh = self.g.h(n) - (64 if s.get("lede") else 0)
        width = self.g.w(8); body = self.th.styles["body"]; items = val["items"]
        size = body["pt"]
        while True:
            if self.list_height(items, size, width) <= boxh:
                return [dict(s, _listPt=size)]
            if size - 2 < body["minPt"]: break
            size -= 2
        # split at original size into fewest balanced slides
        groups = []
        for it in items:
            if isinstance(it, tuple) and it[1] > 1 and groups: groups[-1].append(it)
            else: groups.append([it])
        for pages in range(2, 20):
            per = math.ceil(len(groups) / pages)
            chunks = [sum(groups[i:i + per], []) for i in range(0, len(groups), per)]
            probe = dict(s, _suffix=continued(2, len(chunks))); probe.pop("subtitle", None)
            if all(self.list_height(c, body["pt"], width) <= self.g.h(self.content_zone(probe)[1]) for c in chunks): break
        out, start = [], 1
        for i, c in enumerate(chunks):
            d = dict(s, _listPt=body["pt"], list=dict(val, items=c, start=start))
            if i:
                d["_suffix"] = continued(i + 1, len(chunks)); d.pop("subtitle", None); d.pop("lede", None)
            out.append(d); start += len(c)
        return out

    def split_table(self, s, val):
        self.title_fit(s)
        rows = val["rows"]; th = self.th
        tb, thd = th.styles["tableBody"], th.styles["tableHeader"]
        row_h = tb["pt"] * tb["lineHeight"] + 2 * (th.cell_pad + 2)
        head_h = thd["pt"] * thd["lineHeight"] + 2 * (th.cell_pad + 3)
        def capacity(d):
            zone = self.g.h(self.content_zone(d)[1]) - (76 if d.get("lede") else 0)
            return int((zone - head_h - 22) // row_h)
        first = capacity(s)
        if val.get("maxRows"): first = val["maxRows"]
        if len(rows) <= first: return [s]
        probe = dict(s, _suffix=continued(2, 3)); probe.pop("subtitle", None); probe.pop("lede", None)
        lim = min(first, capacity(probe)) if not val.get("maxRows") else val["maxRows"]
        pages = math.ceil(len(rows) / lim); base, extra = divmod(len(rows), pages)
        out, i = [], 0
        for p in range(pages):
            n = base + (1 if p < extra else 0)
            d = dict(s, table=dict(val, rows=rows[i:i + n], range=(i + 1, i + n, len(rows))))
            if p: d["_suffix"] = continued(p + 1, pages); d.pop("subtitle", None); d.pop("lede", None)
            out.append(d); i += n
        return out

def continued(i, n): return " (continued)" if n == 2 else " (continued %d/%d)" % (i, n)

def blocks(s):
    b = []
    for k in ("list", "table", "chart", "image", "text"):
        if k in s: b.append((k, s[k]))
    for k in ("columns",):
        for c in s.get(k, []): b.append(("column", c))
    for k, n in (("kpis", "kpi"), ("cards", "card"), ("steps", "step"), ("locations", "location")):
        if s.get(k): b.append((n, s[k]))
    for extra in s.get("blocks", []): b.append(extra)
    return b

# ---------------------------------------------------------------- renderers

def render(deck, s, number):
    th, g, spec = deck.th, deck.g, deck.spec
    lay = s["_layout"]; L = spec["layouts"][lay]
    bg = th.color(L["background"]) if L["background"] != "palette.background" else th.p["background"]
    sl = Slide(th, g, lay, bg)
    R = globals()["r_" + lay]
    R(deck, sl, s, L)
    footer(deck, sl, s, lay, number)
    return sl

def footer(deck, sl, s, lay, number):
    vis = deck.spec["chrome"]["footer"]["visibility"].get(lay, "both")
    if vis.startswith("none"): return
    g, th = deck.g, deck.th
    col = None
    if lay in ("Section",): col = th.color("onInkMuted")
    if lay == "ImageFull": col = rgba("#FFFFFF", .8)
    if vis.startswith("both"):
        sl.text(g.x(1), g.footer_y, g.w(8), g.footer_h, "footer", deck.footer, anchor="bottom", color=col)
    sl.text(g.x(9), g.footer_y, g.w(4), g.footer_h, "footer", str(number), anchor="bottom", align="right", color=col)

def header(deck, sl, s):
    g, th = deck.g, deck.th
    eyebrow = s.get("eyebrow") or deck.section
    geo = deck.header_geom(s)
    if s.get("_suffix") and s.get("_suffixPlace") == "eyebrow":
        eyebrow = ((eyebrow + " \u00B7 ") if eyebrow else "") + s["_suffix"].strip(" ()")
    x, y, w, h = g.box(1, 8, 1, 1)
    sl.box(x, y - 12, 28, 3, "", "background:%s" % th.p["accent1"])
    if eyebrow: sl.text(x, y, w, h, "eyebrow", eyebrow)
    t = E(s.get("title", ""))
    if s.get("_suffix") and s.get("_suffixPlace") == "title":
        t += '<span style="white-space:nowrap;font-family:\'%s\';font-size:%gpt;font-weight:400;letter-spacing:0;color:%s">%s</span>' % (th.body_font, th.styles["subtitle"]["pt"], th.p["textMuted"], E(s["_suffix"]))
    sl.text(g.x(1), geo["titleY"], g.w(10), geo["titleLines"] * geo["titlePt"] + 8, "title", t, size=geo["titlePt"], raw=True, extra="text-wrap:balance;")
    if s.get("subtitle"):
        sl.text(g.x(1), geo["subtitleY"], g.w(9), 60, "subtitle", s["subtitle"], extra="text-wrap:balance;")
    return geo["r0"], geo["n"]

def stack(sl, x, y, w, h, items, anchor="bottom"):
    inner = "".join('<div style="%s">%s</div>' % (css, content) for css, content in items)
    jc = "flex-end" if anchor == "bottom" else "flex-start"
    sl.box(x, y, w, h, inner, "display:flex;flex-direction:column;justify-content:%s;" % jc)

def r_Title(deck, sl, s, L):
    g, th = deck.g, deck.th
    x, y, w, h = g.box(1, 8, 1, 1)
    sl.text(x, y, w, h, "eyebrow", s.get("brand", deck.footer), color=th.color("onAccentMuted"))
    x, y, w, h = g.box(1, 9, 3, 8)
    items = []
    if s.get("eyebrow"): items.append((sl.style_css("eyebrow", th.color("onAccentMuted"), extra="margin-bottom:18pt"), E(s["eyebrow"])))
    items.append((sl.style_css("display", th.on_accent, extra="margin-bottom:22pt;text-wrap:balance"), E(s["title"])))
    if s.get("subtitle"): items.append((sl.style_css("subtitle", th.color("onAccentMuted"), extra="max-width:%gpt;text-wrap:balance" % g.w(7)), E(s["subtitle"])))
    stack(sl, x, y, w, h, items)
    vx = g.x(10) - 12; sl.box(vx, g.y(3), 1, g.y(11) - g.y(3) - 12, "", "background:%s" % rgba(th.on_accent, .28))
    x, y, w, h = g.box(10, 3, 9, 2)
    sl.text(x, y, w, h, "caption", "<br>".join(E(m) for m in s.get("meta", [])), anchor="bottom", color=th.color("onAccentMuted"), raw=True, size=th.styles["caption"]["pt"] + 1)

def r_Section(deck, sl, s, L):
    g, th = deck.g, deck.th
    deck.section_no += 1; deck.section = s["title"]
    num = "%02d" % deck.section_no
    x, y, w, h = g.box(8, 5, 1, 12)
    sl.text(x, y, w, h + 28 + 52, "sectionNumeral", num, anchor="bottom", align="right", color=rgba(th.on_ink, .09))
    x, y, w, h = g.box(1, 8, 4, 7)
    eb_col = th.p["accent1"] if contrast(th.p["accent1"], th.ink) >= 3 else th.color("onInkMuted")
    items = [(sl.style_css("eyebrow", eb_col, extra="margin-bottom:16pt"), E(s.get("eyebrow") or "Section " + num)),
             (sl.style_css("sectionTitle", th.on_ink, extra="margin-bottom:18pt"), E(s["title"]))]
    if s.get("subtitle"): items.append((sl.style_css("subtitle", th.color("onInkMuted"), extra="max-width:%gpt;text-wrap:balance" % g.w(8)), E(s["subtitle"])))
    stack(sl, x, y, w, h, items)

def list_html(deck, items, size, start=1, numbered=False, color=None, width=None):
    th = deck.th; body = th.styles["body"]; gap = 8 * size / body["pt"]
    out = []
    n = start
    for it in items:
        txt, lvl = (it if isinstance(it, tuple) else (it, 1))
        sz = size if lvl == 1 else round(size * 0.86)
        col = color or (th.p["text"] if lvl == 1 else th.p["textMuted"])
        if numbered and lvl == 1:
            mark = '<span style="position:absolute;left:0;top:0;font-weight:700;color:%s;font-size:%gpt">%02d</span>' % (th.p["accent1"], sz * 0.78, n); n += 1
            ind = 34
        elif lvl == 1:
            mark = '<span style="position:absolute;left:0;top:%gpt;width:%gpt;height:%gpt;background:%s"></span>' % (sz * 0.52, sz * 0.3, sz * 0.3, th.p["accent1"]); ind = 18
        else:
            mark = '<span style="position:absolute;left:22pt;top:0;color:%s">–</span>' % th.hairline; ind = 40
        out.append('<li style="position:relative;padding-left:%gpt;margin-bottom:%gpt;font-size:%gpt;color:%s">%s%s</li>' % (ind, gap, sz, col, mark, E(txt)))
    return '<ul style="%s">%s</ul>' % (deck_style(deck, "body", size), "".join(out))

def deck_style(deck, key, size=None):
    s = deck.th.styles[key]; fam = deck.th.heading_font if s["font"] == "fonts.heading" else deck.th.body_font
    return "font-family:'%s',Aptos,Arial,sans-serif;font-size:%gpt;line-height:%s;font-weight:%d" % (fam, size or s["pt"], s["lineHeight"], s["weight"])

def content(deck, sl, kind, val, x, y, w, h, s):
    th = deck.th
    if kind == "list":
        size = s.get("_listPt", th.styles["body"]["pt"])
        sl.box(x, y, w, h, list_html(deck, val["items"], size, val.get("start", 1), val.get("numbered")))
    elif kind == "text":
        sl.box(x, y, w, h, '<div style="%s;color:%s">%s</div>' % (deck_style(deck, "body"), th.p["text"], "".join("<p style='margin:0 0 8pt'>%s</p>" % E(p) for p in val.split("\n"))))
    elif kind == "table":
        sl.box(x, y, w, h, table_html(deck, val, w))
    elif kind == "chart":
        sl.box(x, y, w, h, chart_svg(deck, val, w, h))
    elif kind == "image":
        sl.box(x, y, w, h, '<img src="%s" style="width:100%%;height:100%%;object-fit:cover">' % val["src"])
    elif kind == "column":
        sub = blocks(val["content"])
        k2, v2 = sub[0]
        content(deck, sl, k2, v2, x, y, w, h, val["content"])

def r_TitleAndContent(deck, sl, s, L):
    g = deck.g
    r0, n = header(deck, sl, s)
    kind, val = blocks(s)[0]
    x, y, w, h = g.box(1, 12, r0, n)
    if s.get("lede"):
        sl.text(x, y, g.w(8), 64, "lede", s["lede"]); y += 76; h -= 76
    if kind in ("list", "text"): w = g.w(8)
    content(deck, sl, kind, val, x, y, w, h, s)

def r_TwoContent(deck, sl, s, L):
    g = deck.g
    r0, n = header(deck, sl, s)
    b = blocks(s)
    kinds = [k if k != "column" else blocks(v["content"])[0][0] for k, v in b]
    texty = [k in ("list", "text") for k in kinds]
    if texty[0] != texty[1]:
        cols = [(1, 4), (6, 7)] if texty[0] else [(1, 7), (9, 4)]
    else:
        cols = [(1, 6), (7, 6)]
    for (k, v), (c, sp) in zip(b, cols):
        x, y, w, h = g.box(c, sp, r0, n)
        content(deck, sl, k, v, x, y, w, h, s)

def r_Comparison(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, n = header(deck, sl, s)
    for (k, v), c, acc in zip(blocks(s), (1, 7), (th.p["accent1"], th.p["accent2"])):
        x, y, w, h = g.box(c, 6, r0, 1)
        sl.text(x, y - 8, w, h, "heading", v["title"], anchor="bottom")
        sl.box(x, y + h + 6, w, 3, "", "background:%s" % acc)
        x, y, w, h = g.box(c, 6, r0 + 1, n - 1)
        content(deck, sl, "column", v, x, y + 12, w, h - 12, s)

def tiles_geom(g, n, spans, x0=None, total=None):
    if spans == "fifths":
        tw = (g.w(12) - 4 * g.g) / 5; return [(g.x(1) + i * (tw + g.g), tw) for i in range(n)]
    out, c = [], 1
    for sp in spans: out.append((g.x(c), g.w(sp))); c += sp
    return out

def change_html(th, change, color=None, size=None, mt=12):
    if not change: return ""
    if color: col = th.p[color.lower()] if color.lower() in th.p else color
    elif change.startswith("+"): col = th.p["success"]
    elif change[0] in "-−": col = th.p["danger"]
    else: col = th.p["textMuted"]
    glyph = "▲ " if change.startswith("+") else ("▼ " if change[0] in "-−" else "")
    return '<div style="font-family:\'%s\';font-size:%gpt;font-weight:700;color:%s;margin-top:%gpt">%s%s</div>' % (th.body_font, size or th.styles["caption"]["pt"] + 1, col, mt, glyph, E(change))

def r_Kpi(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, nrows = header(deck, sl, s)
    k = s["kpis"]; n = len(k)
    spec = L["tiling"][str(n)]
    rowsets = [k[:3], k[3:]] if spec.get("rows") == 2 else [k]
    if spec.get("rows") == 2:
        ys = [(g.y(r0), 140), (g.y(r0) + 140 + 26, 140)]
    else:
        ys = [(g.y(r0 + 1), g.h(5))]
    for (ty, th_), items in zip(ys, rowsets):
        geo = tiles_geom(g, len(items), spec["spans"] if spec.get("rows") != 2 else [4, 4, 4])
        for i, ((x, w), kp) in enumerate(zip(geo, items)):
            first = (items is rowsets[0] and i == 0)
            vs = th.styles["kpiValue"]["pt"] * (spec.get("valueStyleFactor", 1))
            while text_w(kp["value"], vs, -0.02) > w and vs > th.styles["kpiValue"]["minPt"]: vs -= 2
            inner = '<div style="height:2pt;background:%s;margin-bottom:18pt"></div>' % (th.p["accent1"] if first else th.hairline)
            inner += '<div style="%s">%s</div>' % (sl.style_css("kpiValue", size=vs), E(kp["value"]))
            inner += '<div style="%s;margin-top:10pt;max-width:%gpt">%s</div>' % (sl.style_css("kpiLabel"), w - 8, E(kp["label"]))
            inner += change_html(th, kp.get("change"), kp.get("color"))
            sl.box(x, ty, w, th_, inner)
    if s.get("takeaway"):
        x, y, w, h = g.box(1, 8, 11, 2)
        sl.box(x, y - 14, g.w(12), 1, "", "background:%s" % th.hairline)
        sl.text(x, y, w, h, "lede", s["takeaway"])

def r_Cards(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, n = header(deck, sl, s)
    cards = s["cards"]; spec = L["tiling"][str(len(cards))]
    rows = spec.get("rows") or [spec["spans"]]
    zy, zh = g.y(r0), g.h(n)
    rh = zh if len(rows) == 1 else (zh - 24) / 2
    idx = 0
    def group_h(c, w):
        iw = w - 40; hd = th.styles["heading"]; bs = th.styles["bodySmall"]
        h = len(wrap(c["title"], hd["pt"] * th.hfac, iw)) * hd["pt"] * hd["lineHeight"]
        if c.get("text"): h += 8 + len(wrap(c["text"], bs["pt"] * th.bfac, iw)) * bs["pt"] * bs["lineHeight"]
        if c.get("items"): h += 10 + len(c["items"]) * (bs["pt"] * bs["lineHeight"] + 3)
        return h
    for ri, spans in enumerate(rows):
        geo = tiles_geom(g, len(spans), spans)
        shared = max(group_h(cards[idx + j], w) for j, (x, w) in enumerate(geo))
        for x, w in geo:
            c = cards[idx]; idx += 1
            y = zy + ri * (rh + 24)
            acc = th.p.get((c.get("color") or "accent1").lower(), th.p["accent1"])
            top = '<div style="%s">%02d</div>' % (sl.style_css("eyebrow", th.p["textMuted"]), idx)
            bot = '<div style="%s">%s</div>' % (sl.style_css("heading"), E(c["title"]))
            if c.get("text"): bot += '<div style="%s;margin-top:8pt">%s</div>' % (sl.style_css("bodySmall"), E(c["text"]))
            if c.get("items"):
                bot += '<div style="margin-top:10pt">' + "".join('<div style="%s;padding-left:16pt;position:relative;margin-top:3pt"><span style="position:absolute;left:0;color:%s">–</span>%s</div>' % (sl.style_css("bodySmall", th.p["text"]), acc, E(i)) for i in c["items"]) + "</div>"
            inner = '<div style="position:absolute;left:0;top:0;right:0;height:3pt;background:%s"></div>' % acc
            inner += '<div style="position:absolute;left:20pt;right:20pt;top:23pt">%s</div><div style="position:absolute;left:20pt;right:20pt;top:%gpt">%s</div>' % (top, rh - 20 - shared, bot)
            sl.box(x, y, w, rh, inner, "background:%s" % th.surface)

def r_Process(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, nrows = header(deck, sl, s)
    steps = s["steps"]; spec = L["tiling"][str(len(steps))]
    geo = tiles_geom(g, len(steps), spec["spans"])
    y0 = g.y(r0 + 1); d = 40
    for i, ((x, w), st) in enumerate(zip(geo, steps)):
        first = i == 0
        if i < len(steps) - 1:
            nx = geo[i + 1][0]
            sl.box(x + d + 8, y0 + d / 2 - 1, nx - x - d - 16, 2, "", "background:%s" % th.hairline)
        num = st.get("number") or "%02d" % (i + 1)
        sl.box(x, y0, d, d, '<div style="width:100%%;height:100%%;display:flex;align-items:center;justify-content:center;%s">%s</div>' % (sl.style_css("heading", th.on_accent if first else th.p["accent1"], 15), E(num)),
               "border-radius:50%%;border:2pt solid %s;background:%s" % (th.p["accent1"], th.p["accent1"] if first else th.p["background"]))
        inner = '<div style="%s">%s</div>' % (sl.style_css("heading"), E(st["title"]))
        if st.get("text"): inner += '<div style="%s;margin-top:8pt;padding-right:8pt">%s</div>' % (sl.style_css("bodySmall"), E(st["text"]))
        sl.box(x, y0 + d + 22, w, g.h(6) - d - 22, inner)
    if s.get("takeaway"):
        x, y, w, h = g.box(1, 8, 11, 2)
        sl.box(x, y - 14, g.w(12), 1, "", "background:%s" % th.hairline)
        sl.text(x, y, w, h, "lede", s["takeaway"])

def r_Quote(deck, sl, s, L):
    g, th = deck.g, deck.th
    x, y, w, h = g.box(1, 1, 2, 3)
    sl.text(x - 4, y - 10, w + 40, h + 40, "display", "“", color=th.p["accent1"], size=120)
    x, y, w, h = g.box(2, 9, 4, 6)
    q = s["text"]; size = th.styles["quote"]["pt"]
    while len(wrap(q, size, w)) > 5 and size > th.styles["quote"]["minPt"]: size -= 2
    sl.text(x, y, w, h, "quote", q, size=size)
    x, y, w, h = g.box(2, 7, 10, 2)
    sl.box(x, y - 14, 48, 2, "", "background:%s" % th.p["accent1"])
    parts = s.get("subtitle", "").split(", ", 1)
    att = '<span style="color:%s;font-weight:700">%s</span>' % (th.p["text"], E(parts[0])) + ((", " + E(parts[1])) if len(parts) > 1 else "")
    sl.text(x, y, w, h, "bodySmall", att, raw=True)

def r_ImageFull(deck, sl, s, L):
    g, th = deck.g, deck.th
    sl.box(0, 0, g.W, g.H, '<img src="%s" style="width:100%%;height:100%%;object-fit:cover;object-position:%s">' % (s["image"]["src"], s["image"].get("focus", "center")))
    sl.box(0, 0, g.W, g.H, "", "background:linear-gradient(180deg,%s 35%%,%s 100%%),linear-gradient(90deg,%s 0%%,%s 62%%)" % (rgba(th.ink, 0), rgba(th.ink, .78), rgba(th.ink, .4), rgba(th.ink, 0)))
    x, y, w, h = g.box(1, 8, 7, 5)
    items = []
    if s.get("eyebrow"): items.append((sl.style_css("eyebrow", rgba("#FFFFFF", .75), extra="margin-bottom:12pt"), E(s["eyebrow"])))
    items.append((sl.style_css("title", "#FFFFFF", round_size(th.t["typeScale"]["heading1"] * 1.5), "margin-bottom:12pt"), E(s["title"])))
    if s.get("subtitle"): items.append((sl.style_css("subtitle", rgba("#FFFFFF", .8), extra="text-wrap:balance"), E(s["subtitle"])))
    stack(sl, x, y, w, h, items)

def r_Closing(deck, sl, s, L):
    g, th = deck.g, deck.th
    x, y, w, h = g.box(1, 6, 3, 8)
    items = []
    if s.get("eyebrow"): items.append((sl.style_css("eyebrow", th.color("onAccentMuted"), extra="margin-bottom:18pt"), E(s["eyebrow"])))
    items.append((sl.style_css("display", th.on_accent, round_size(th.t["typeScale"]["title"] * 1.5), "margin-bottom:20pt;text-wrap:balance"), E(s["title"])))
    if s.get("subtitle"): items.append((sl.style_css("subtitle", th.color("onAccentMuted"), extra="text-wrap:balance"), E(s["subtitle"])))
    stack(sl, x, y, w, h, items)
    if s.get("list"):
        x, y, w, h = g.box(8, 5, 3, 8)
        li = "".join('<div style="border-top:1pt solid %s;padding-top:12pt;margin-top:18pt"><div style="%s">%02d</div><div style="%s;margin-top:6pt">%s</div></div>' % (rgba(th.on_accent, .28), sl.style_css("eyebrow", th.color("onAccentMuted")), i + 1, sl.style_css("body", th.on_accent), E(t)) for i, t in enumerate(s["list"]["items"]))
        sl.box(x, y, w, h, li, "display:flex;flex-direction:column;justify-content:flex-end")

def r_CaseStudy(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, n = header(deck, sl, s)
    x, y, w, h = g.box(1, 7, r0, n)
    parts = []
    for i, c in enumerate(s["cards"]):
        sep = "" if i == 0 else '<div style="height:1pt;background:%s;margin:16pt 0"></div>' % th.hairline
        parts.append(sep + '<div style="%s">%s</div><div style="%s;margin-top:6pt">%s</div>' % (sl.style_css("eyebrow"), E(c["title"]), sl.style_css("body"), E(c.get("text") or "; ".join(c.get("items", [])))))
    sl.box(x, y, w, h, "".join(parts))
    x, y, w, h = g.box(9, 4, r0, n)
    vs = round_size(th.styles["kpiValue"]["pt"] * 0.67)
    nk = len(s["kpis"]); lab = th.styles["kpiLabel"]["pt"] * th.styles["kpiLabel"]["lineHeight"]
    fit = (h - 3 - 40 - (nk - 1) * 26) / nk - (4 + lab + 6 + 14)
    vs = max(th.styles["kpiValue"]["minPt"] * 0.75, min(vs, math.floor(fit)))
    kp = "".join('<div style="%s"><div style="%s">%s</div><div style="%s;margin-top:4pt">%s</div>%s</div>' % ("" if i == 0 else "border-top:1pt solid %s;padding-top:12pt" % th.hairline, sl.style_css("kpiValue", size=vs), E(k["value"]), sl.style_css("kpiLabel"), E(k["label"]), change_html(th, k.get("change"), k.get("color"), mt=6)) for i, k in enumerate(s["kpis"]))
    sl.box(x, y, w, h, '<div style="height:3pt;background:%s"></div><div style="box-sizing:border-box;height:calc(100%% - 3pt);padding:20pt 24pt;display:flex;flex-direction:column;justify-content:space-between">%s</div>' % (th.p["accent1"], kp), "background:%s" % th.surface)

def r_Logos(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, nr = header(deck, sl, s)
    cards = s["cards"]; n = len(cards)
    per = 4 if n <= 8 else 6; span = 12 // per
    rows = [cards[i:i + per] for i in range(0, n, per)]
    th_ = 96; total = len(rows) * th_ + (len(rows) - 1) * 24
    y0 = g.y(r0) + (g.h(nr) - total) / 2 - 10
    for ri, row in enumerate(rows):
        for ci, c in enumerate(row):
            x, w = g.x(1 + ci * span), g.w(span)
            sl.box(x, y0 + ri * (th_ + 24), w, th_, '<img src="%s" alt="%s" style="max-width:62%%;max-height:44%%;filter:grayscale(1);opacity:.8">' % (c["image"], E(c["title"])),
                   "border:1pt solid %s;display:flex;align-items:center;justify-content:center;background:%s" % (th.hairline, th.p["background"]))
    if s.get("takeaway"):
        x, y, w, h = g.box(1, 8, 12, 1); sl.text(x, y, w, h, "bodySmall", s["takeaway"])

def r_Coverage(deck, sl, s, L):
    g, th = deck.g, deck.th
    r0, n = header(deck, sl, s)
    x, y, w, h = g.box(1, 8, r0, n)
    dots = "radial-gradient(%s 1.1pt, transparent 1.3pt) 0 0/12pt 12pt" % th.hairline
    sl.box(x, y, w, h, '<div style="position:absolute;left:20pt;bottom:16pt;%s">%s</div>' % (sl.style_css("caption"), E(s.get("mapLabel", ""))), "background:%s, %s" % (dots, th.surface))
    for i, loc in enumerate(s["locations"]):
        px, py = x + loc["left"] * w, y + loc["top"] * h
        sl.box(px - 11, py - 11, 22, 22, '<div style="width:100%%;height:100%%;display:flex;align-items:center;justify-content:center;font:700 10pt \'%s\';color:%s">%d</div>' % (th.body_font, th.on_accent, i + 1),
               "border-radius:50%%;background:%s;border:2pt solid %s;box-shadow:0 0 0 6pt %s" % (th.p["accent1"], th.p["background"], rgba(th.p["accent1"], .14)))
    x, y, w, h = g.box(9, 4, r0, n)
    items = "".join('<div style="display:flex;gap:10pt;margin-bottom:16pt"><div style="%s;flex:none;width:18pt">%02d</div><div><div style="%s">%s</div><div style="%s;margin-top:2pt">%s</div></div></div>' % (
        sl.style_css("caption", th.p["accent1"], extra="font-weight:700;padding-top:3pt"), i + 1, sl.style_css("heading", size=16), E(l["name"]), sl.style_css("bodySmall"), E(l.get("detail", ""))) for i, l in enumerate(s["locations"][:7]))
    sl.box(x, y, w, h, items)

# ---------------------------------------------------------------- table and chart

def is_num(v): return isinstance(v, (int, float)) or bool(re.fullmatch(r"[\-−+]?[\d,.]+ ?%?", str(v)))

def table_html(deck, t, width):
    th = deck.th; cols = t["columns"]; pad = th.cell_pad
    hs = sl_style = deck_style(deck, "tableHeader"); bs = deck_style(deck, "tableBody")
    num = [all(is_num(r[i]) for r in t["rows"]) for i in range(len(cols))]
    rule = "border-bottom:%gpt solid %s;" % (th.header_rule[1], th.header_rule[0]) if th.header_rule else ""
    head = "".join('<th style="%s;text-align:%s;padding:%gpt %gpt;background:%s;color:%s;%s">%s</th>' % (hs, "right" if num[i] else "left", pad + 3, pad + 5, th.header_fill, th.header_text, rule, E(c)) for i, c in enumerate(cols))
    body = []
    for ri, r in enumerate(t["rows"]):
        bgc = th.surface if ri % 2 else th.p["background"]
        cells = "".join('<td style="%s;text-align:%s;padding:%gpt %gpt;color:%s;border-bottom:0.5pt solid %s">%s</td>' % (bs, "right" if num[i] else "left", pad + 2, pad + 5, cell_color(th, v), th.hairline, E(str(v))) for i, v in enumerate(r))
        body.append('<tr style="background:%s">%s</tr>' % (bgc, cells))
    cap = ""
    if t.get("range"):
        a, b, n = t["range"]
        cap = '<div style="%s;color:%s;text-align:right;margin-top:8pt">Rows %d–%d of %d</div>' % (deck_style(deck, "caption"), th.p["textMuted"], a, b, n)
    return "<table><thead><tr>%s</tr></thead><tbody>%s</tbody></table>%s" % (head, "".join(body), cap)

def cell_color(th, v):
    return {"Outage": th.p["danger"], "Degraded": th.p["warning"]}.get(str(v), th.p["text"])

def chart_svg(deck, c, W, H):
    th = deck.th; f = th.body_font
    series = c["series"]; cats = c["categories"]; typ = c.get("type", "Column")
    title_h = 30 if c.get("title") else 0
    cols = [th.p["accent%d" % (i + 1)] for i in range(6)]
    out = ['<svg width="%gpt" height="%gpt" viewBox="0 0 %g %g" xmlns="http://www.w3.org/2000/svg" style="overflow:visible;font-family:\'%s\',Aptos,Arial">' % (W, H, W, H, f)]
    if c.get("title"):
        out.append('<text x="0" y="14" font-size="15" font-weight="600" fill="%s">%s</text>' % (th.p["text"], E(c["title"])))
        if len(series) > 1:
            lx = W
            for i, (nm, _) in reversed(list(enumerate(series))):
                tw = text_w(nm, 11) + 20; lx -= tw
                out.append('<rect x="%g" y="5" width="9" height="9" fill="%s"/><text x="%g" y="13.5" font-size="11" fill="%s">%s</text>' % (lx, cols[i], lx + 13, th.p["textMuted"], E(nm)))
    allv = [v for _, vs in series for v in vs]
    vmax = max(allv); vmin = min(0, min(allv))
    if c.get("min") is not None: vmin = c["min"]
    step = nice_step((vmax - vmin) / 4); top = math.ceil(vmax * (1.08 if c.get("labels") else 1) / step) * step
    if typ == "Bar":
        lw = max(text_w(x, 11) for x in cats) + 12
        px, pw, py, ph = lw, W - lw - 40, title_h + 14, H - title_h - 34
        bh = ph / len(cats)
        for k in range(0, int(top / step) + 1):
            gx = px + pw * k * step / top
            out.append('<line x1="%g" x2="%g" y1="%g" y2="%g" stroke="%s" stroke-width="0.75"/>' % (gx, gx, py, py + ph, th.hairline))
            out.append('<text x="%g" y="%g" font-size="11" fill="%s" text-anchor="middle">%s</text>' % (gx, py + ph + 16, th.p["textMuted"], fmt_num(k * step)))
        for i, cat in enumerate(cats):
            v = series[0][1][i]; bw = pw * v / top; y = py + i * bh + bh * 0.18
            col = cols[0] if cat != "Other" and not cat.startswith("Other") else th.hairline
            out.append('<rect x="%g" y="%g" width="%g" height="%g" fill="%s"/>' % (px, y, bw, bh * 0.64, col))
            out.append('<text x="%g" y="%g" font-size="11" fill="%s" text-anchor="end">%s</text>' % (px - 10, y + bh * 0.32 + 4, th.p["text"], E(cat)))
            out.append('<text x="%g" y="%g" font-size="11" fill="%s">%s</text>' % (px + bw + 6, y + bh * 0.32 + 4, th.p["textMuted"], fmt_num(v)))
        out.append("</svg>"); return "".join(out)
    px, py = 34, title_h + 16; pw, ph = W - px - 4, H - py - 26
    rng = top - vmin
    for k in range(0, int(round(rng / step)) + 1):
        v = vmin + k * step; gy = py + ph - ph * (v - vmin) / rng
        out.append('<line x1="%g" x2="%g" y1="%g" y2="%g" stroke="%s" stroke-width="%s"/>' % (px, px + pw, gy, gy, th.hairline, "1" if v == vmin else "0.75"))
        out.append('<text x="%g" y="%g" font-size="11" fill="%s" text-anchor="end">%s</text>' % (px - 8, gy + 4, th.p["textMuted"], fmt_num(v)))
    n = len(cats); slot = pw / n
    for i, cat in enumerate(cats):
        out.append('<text x="%g" y="%g" font-size="11" fill="%s" text-anchor="middle">%s</text>' % (px + slot * (i + .5), py + ph + 18, th.p["textMuted"], E(cat)))
    if typ == "Column":
        ns = len(series); gw = slot * 0.62; bw = gw / ns
        for si, (nm, vs) in enumerate(series):
            for i, v in enumerate(vs):
                bh = ph * (v - vmin) / rng
                out.append('<rect x="%g" y="%g" width="%g" height="%g" fill="%s"/>' % (px + slot * i + (slot - gw) / 2 + si * bw + (1 if ns > 1 else 0), py + ph - bh, bw - (2 if ns > 1 else 0), bh, cols[si] if ns > 1 or c.get("highlightLast") is None or i == n - 1 else rgba(cols[si], .35)))
                if c.get("labels"):
                    out.append('<text x="%g" y="%g" font-size="11" font-weight="600" fill="%s" text-anchor="middle">%s</text>' % (px + slot * (i + .5), py + ph - bh - 6, th.p["text"], fmt_num(v)))
    else:
        for si, (nm, vs) in enumerate(series):
            pts = [(px + slot * (i + .5), py + ph - ph * (v - vmin) / rng) for i, v in enumerate(vs)]
            out.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="2.5"/>' % (" ".join("%g,%g" % p for p in pts), cols[si]))
            for i, (a, b) in enumerate(pts):
                out.append('<circle cx="%g" cy="%g" r="4" fill="%s" stroke="%s" stroke-width="2"/>' % (a, b, th.p["background"], cols[si]))
                if c.get("labels"):
                    out.append('<text x="%g" y="%g" font-size="11" font-weight="600" fill="%s" text-anchor="middle">%s</text>' % (a, b - 10, th.p["text"], fmt_num(vs[i]) + c.get("suffix", "")))
    out.append("</svg>"); return "".join(out)

def nice_step(x):
    e = 10 ** math.floor(math.log10(x)); m = x / e
    return (1 if m <= 1 else 2 if m <= 2 else 2.5 if m <= 2.5 else 5 if m <= 5 else 10) * e

def fmt_num(v): return ("%d" % v) if float(v).is_integer() else ("%.1f" % v)

# ---------------------------------------------------------------- sample content

def logo_svg(out, name, i):
    shapes = ['<circle cx="14" cy="16" r="10" fill="#333"/>', '<rect x="4" y="6" width="20" height="20" fill="#333"/>', '<path d="M4 26 L14 6 L24 26 Z" fill="#333"/>',
              '<rect x="4" y="6" width="20" height="20" rx="6" fill="none" stroke="#333" stroke-width="4"/>']
    p = out / "assets" / ("logo-%s.svg" % name.lower())
    w = 40 + len(name) * 13
    p.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="32" viewBox="0 0 %d 32">%s<text x="34" y="23" font-family="Aptos,Arial" font-size="20" font-weight="%d" fill="#333" letter-spacing="%s">%s</text></svg>' % (w, w, shapes[i % 4], 700 if i % 2 else 500, "0.5" if i % 3 == 0 else "0", name))
    return str(p)

def samples(out):
    img = str(out / "assets" / "Sonoma.jpg")
    q = ["Q1", "Q2", "Q3", "Q4"]
    L = {}
    L["Title"] = {"title": "Quarterly business review", "eyebrow": "Third quarter 2026", "subtitle": "Performance, decisions and priorities for the fourth quarter", "meta": ["Prepared by Finance & Strategy", "26 October 2026"], "brand": "Northwind Traders"}
    L["Section"] = {"title": "Performance", "subtitle": "Revenue, margin and customer growth against plan"}
    L["TitleAndContent"] = {"title": "What changed this quarter", "list": {"items": [
        "Enterprise renewals closed two weeks ahead of the quarter end", ("Three of four strategic accounts expanded their contracts", 2),
        "Self-service onboarding now covers 70% of new customers", "Support cost per ticket fell for the third quarter in a row",
        "The EMEA pricing change is live in all markets", ("Early churn is flat; we review it again in November", 2)]}}
    L["TitleAndContent-table"] = {"title": "Next-quarter priorities", "subtitle": "Owners confirmed at the October leadership meeting", "table": {"columns": ["Priority", "Owner", "Due", "Budget (USD k)", "Status"], "rows": [
        ["Launch partner portal", "Sales operations", "15 Nov", "420", "On track"], ["Consolidate billing systems", "Finance", "30 Nov", "310", "At risk"],
        ["Hire two solution architects", "Engineering", "15 Dec", "280", "On track"], ["Expand APAC support hours", "Customer success", "31 Dec", "190", "Planned"],
        ["Refresh the security review", "IT", "31 Dec", "75", "On track"]]}}
    L["TitleAndContent-chart"] = {"title": "Revenue grew in every quarter", "lede": "Revenue reached USD 5.7m in Q4, 36% above Q1, with the strongest growth in enterprise renewals.",
        "chart": {"type": "Column", "title": "Revenue (USD m)", "categories": q, "series": [("Revenue", [4.2, 4.8, 5.1, 5.7])], "labels": True, "highlightLast": True}}
    L["TwoContent"] = {"title": "Performance trend", "subtitle": "Revenue grew 36% from Q1; margin improved five points",
        "chart": {"type": "Column", "title": "Revenue (USD m)", "categories": q, "series": [("Revenue", [4.2, 4.8, 5.1, 5.7])], "labels": True},
        "blocks": [("chart", {"type": "Line", "title": "Operating margin (%)", "categories": q, "series": [("Margin", [31, 33, 35, 36])], "labels": True, "min": 25, "suffix": "%"})]}
    L["TwoContent-textVisual"] = {"title": "Quarterly ticket demand", "eyebrow": "Workplace support",
        "list": {"items": ["Tickets fell 10% to 1,080", "East depot fell most, by 70 tickets", "Remote staff is the only area that grew", "Head office demand is stable"]},
        "chart": {"type": "Column", "title": "Tickets by location", "categories": ["Head office", "East depot", "West depot", "Remote"], "series": [("Previous quarter", [380, 418, 262, 140]), ("Current quarter", [342, 348, 234, 156])]}}
    L["Comparison"] = {"title": "Service desk today and in the pilot", "eyebrow": "Workplace support", "columns": [
        {"title": "Today", "content": {"list": {"items": ["Four intake routes: email, phone, chat and walk-up", "Tickets are triaged by whoever is free", "No agreed owner for remote staff", "Monthly reporting from spreadsheets"]}}},
        {"title": "Pilot", "content": {"list": {"items": ["One intake form with published categories", "A named triage owner on every shift", "Remote staff handled by the service desk", "A weekly dashboard from the ticket system"]}}}]}
    L["Kpi"] = {"title": "Third quarter at a glance", "kpis": [
        {"value": "$5.7m", "label": "Revenue", "change": "+12% vs Q2"}, {"value": "36%", "label": "Operating margin", "change": "+1 pt vs Q2"},
        {"value": "148", "label": "New customers", "change": "+22% vs Q2"}, {"value": "2.1%", "label": "Monthly churn", "change": "−0.4 pt vs Q2", "color": "success"}],
        "takeaway": "Growth came from renewals and self-service onboarding; churn improved without extra discounting."}
    L["Cards"] = {"title": "Service scope", "subtitle": "Four service areas with explicit ownership and a common intake", "eyebrow": "Workplace support", "cards": [
        {"title": "User incidents", "text": "Faults and how-to questions from every location.", "items": ["Triage", "Ownership", "Escalation"]},
        {"title": "Device lifecycle", "text": "Laptops and phones from order to return.", "items": ["Onboarding", "Replacement", "Returns"]},
        {"title": "Access requests", "text": "Accounts and permissions with an approval trail.", "items": ["Approval", "Least privilege", "Evidence"]},
        {"title": "Service changes", "text": "Planned changes that affect how people work.", "items": ["Notice", "Handover", "Review"]}]}
    L["Process"] = {"title": "A controlled service transition", "subtitle": "Agree the baseline, test the change and review the evidence", "eyebrow": "Workplace support", "steps": [
        {"title": "Baseline", "text": "Agree request categories and reconcile the quarterly intake."}, {"title": "Pilot", "text": "Test the triage queue with a nominated service owner at East depot."},
        {"title": "Review", "text": "Check demand, repeat requests and user feedback before rollout."}, {"title": "Roll out", "text": "Publish ownership, working hours and the escalation route."}],
        "takeaway": "Each step ends with a decision recorded in the service review."}
    L["Quote"] = {"text": "One front door changed the conversation. People stopped asking who to call and started telling us what they needed.", "subtitle": "Maria Lopez, Service desk lead, East depot"}
    L["ImageFull"] = {"title": "East depot goes first", "eyebrow": "Pilot site", "subtitle": "160 people, 348 tickets last quarter and a team that asked to lead", "image": {"src": img, "fullBleed": True, "focus": "center 70%"}}
    L["Closing"] = {"title": "Decisions we need today", "eyebrow": "Next steps", "subtitle": "Owners and dates go into the Q4 plan this week.", "list": {"items": [
        "Approve the partner portal budget", "Confirm the billing consolidation owner", "Agree the APAC support hours"]}}
    L["CaseStudy"] = {"title": "Remote staff: one queue instead of four", "eyebrow": "Workplace support", "cards": [
        {"title": "Challenge", "text": "Remote staff raised requests through four channels, and 1 in 5 tickets was reassigned at least once."},
        {"title": "Approach", "text": "A single intake form, published categories and a named triage owner on every shift."},
        {"title": "Outcome", "text": "Reassignments fell by half within six weeks and first-response time is now under two hours."}],
        "kpis": [{"value": "11%", "label": "Tickets reassigned", "change": "\u221210 pt since July", "color": "success"}, {"value": "1.8 h", "label": "First response", "change": "\u22122.6 h", "color": "success"}, {"value": "4.6/5", "label": "User rating", "change": "+0.7"}]}
    names = ["Northwind", "Contoso", "Fabrikam", "Tailspin", "Wingtip", "Litware", "Adatum", "Proseware"]
    L["Logos"] = {"title": "Accounts that renewed this quarter", "cards": [{"title": n, "image": logo_svg(out, n, i)} for i, n in enumerate(names)]}
    L["Coverage"] = {"title": "Supported locations", "subtitle": "600 people across four service locations", "eyebrow": "Workplace support", "mapLabel": "Service region · schematic", "locations": [
        {"name": "Head office", "left": .30, "top": .38, "detail": "240 people · 342 tickets · workplace team on site"},
        {"name": "East depot", "left": .74, "top": .30, "detail": "160 people · 348 tickets · pilot site"},
        {"name": "West depot", "left": .16, "top": .70, "detail": "120 people · 234 tickets · regional support"},
        {"name": "Remote staff", "left": .58, "top": .72, "detail": "80 people · 156 tickets · service desk"}]}
    return L

def overflow_samples():
    items = ["Choose the right document format for the audience", "Compose content with the PowerShell DSL", "Validate the generated file before sharing it",
             "Keep data and presentation concerns separate", "Use themes instead of per-slide formatting", "Let layouts choose geometry from content",
             "Split long tables instead of shrinking them", "Write speaker notes for every decision slide", "Prefer charts for trends and tables for lookups",
             "Name every image with alternative text", "Check the deck on Windows and macOS", "Keep one idea per slide",
             "Review the rhythm report before publishing", "Archive the source data next to the deck"]
    rows = []
    teams = ["Platform", "Identity", "Network", "Workplace", "Data", "Security"]
    for i in range(30):
        rows.append(["CHG-%04d" % (2140 + i * 7), ["Patch hypervisor cluster", "Rotate signing certificates", "Replace core switch", "Migrate file shares", "Upgrade backup agents", "Retire legacy VPN"][i % 6] + (" (phase %d)" % (i // 6 + 1)),
                     teams[i % 6], "%02d Nov" % (1 + i % 28), ["Approved", "Scheduled", "Completed", "Approved", "Scheduled"][i % 5], "%d" % ((i * 37) % 9 * 15 + 30)])
    return {
        "list": {"title": "Habits that keep generated decks readable", "eyebrow": "Training workshop", "list": {"items": items}},
        "table": {"title": "Change calendar for November", "eyebrow": "Change advisory board", "table": {"columns": ["Change", "Summary", "Team", "Window", "State", "Minutes"], "rows": rows}},
        "title": {"title": "Why the remote staff intake pilot should continue through the first quarter of next year before any wider rollout is approved by the board", "eyebrow": "Workplace support",
                  "list": {"items": ["Demand is still settling after the July change", "Two of three success measures are met", "The third depends on the new ticket categories", "Extending costs nothing beyond current staffing"]}},
    }

def export_sample():
    svc = ["Email", "Identity", "VPN", "File shares", "Printing", "Telephony", "Intranet", "HR portal", "Payroll", "CRM", "ERP", "Service desk",
           "Backup", "Monitoring", "Wi-Fi", "Video meetings", "Mobile devices", "Laptops", "Badge access", "Data warehouse", "Web site", "Source control", "Build servers", "Expenses"]
    owners = ["Workplace", "Identity", "Network", "Workplace", "Workplace", "Network", "Digital", "People", "Finance", "Sales ops", "Finance", "Workplace",
              "Platform", "Platform", "Network", "Workplace", "Workplace", "Workplace", "Facilities", "Data", "Digital", "Engineering", "Engineering", "Finance"]
    status = ["Operational"] * 24; status[2] = "Degraded"; status[13] = "Degraded"; status[4] = "Outage"; status[14] = "Degraded"
    tickets = [212, 188, 164, 97, 143, 45, 22, 38, 19, 61, 74, 130, 12, 28, 118, 84, 69, 156, 31, 17, 9, 26, 33, 41]
    avail = ["99.98%", "99.95%", "98.70%", "99.90%", "97.20%", "99.99%", "99.97%", "99.92%", "99.99%", "99.94%", "99.91%", "99.96%",
             "100.00%", "99.40%", "98.90%", "99.93%", "99.88%", "99.95%", "99.99%", "99.97%", "99.99%", "99.98%", "99.80%", "99.96%"]
    return [{"Service": s, "Owner": o, "Status": st, "Tickets": t, "Availability": a} for s, o, st, t, a in zip(svc, owners, status, tickets, avail)]

def export_deck(spec, rows, title, prop, category, value):
    """Export-OfficePowerPoint deck structure (spec.export)."""
    n = len(rows); slides = []
    slides.append({"title": title, "subtitle": "%d services · status, demand and availability" % n, "meta": ["Exported from Get-ServiceReport", "26 October 2026"], "eyebrow": "Service review", "brand": "IT operations"})
    group = category if category and 2 <= len({r[category] for r in rows}) <= 6 else next(p for p in prop if isinstance(rows[0][p], str) and 2 <= len({r[p] for r in rows}) <= 6)
    counts = {}
    for r in rows: counts[r[group]] = counts.get(r[group], 0) + 1
    tone = {"Outage": "danger", "Degraded": "warning"}
    k = [{"value": str(n), "label": "Services"}]
    for gname, c in sorted(counts.items(), key=lambda x: -x[1])[:3]:
        k.append({"value": str(c), "label": gname, "change": "%d%% of services" % round(100 * c / n), "color": tone.get(gname, "success")})
    if isinstance(rows[0][value], (int, float)):
        k.append({"value": "{:,}".format(sum(r[value] for r in rows)), "label": "Total " + value.lower()})
    small = [(gname, [r[category] for r in rows if r[group] == gname]) for gname, c in sorted(counts.items(), key=lambda x: x[1]) if c <= 3]
    take = " ".join("%s: %s." % (gname, ", ".join(m)) for gname, m in small)
    slides.append({"title": "At a glance", "kpis": k[:5], "layout": "Kpi", "takeaway": take or None})
    slides.append({"title": title, "table": {"columns": prop, "rows": [[r[p] for p in prop] for r in rows]}, "layout": "TitleAndContent"})
    top = sorted(rows, key=lambda r: -r[value])
    cats = [r[category] for r in top[:12]]; vals = [r[value] for r in top[:12]]
    if len(rows) > 12: cats.append("Other (%d)" % (len(rows) - 12)); vals.append(sum(r[value] for r in top[12:]))
    slides.append({"title": "%s by %s" % (value, category.lower()), "chart": {"type": "Bar" if len(cats) > 12 else "Column", "title": "%s this quarter, top 12 %s" % (value, category.lower() + "s"), "categories": cats, "series": [(value, vals)]}, "layout": "TitleAndContent"})
    return slides

# ---------------------------------------------------------------- driver

def write(out, name, sl, title):
    p = out / "html" / (name + ".html"); p.write_text(sl.html(title)); return p

def shoot(html_path, png, attempts=3):
    for i in range(attempts):
        try:
            return _shoot(html_path, png)
        except subprocess.CalledProcessError:
            if i == attempts - 1: raise
            import time; time.sleep(2)

def _shoot(html_path, png):
    import tempfile, shutil
    profile = tempfile.mkdtemp(prefix="pptmock-")
    try:
        _run_edge(html_path, png, profile)
    finally:
        shutil.rmtree(profile, ignore_errors=True)

def _run_edge(html_path, png, profile):
    subprocess.run([EDGE, "--headless", "--disable-gpu", "--no-first-run", "--user-data-dir=%s" % profile, "--hide-scrollbars", "--allow-file-access-from-files", "--force-device-scale-factor=1",
                    "--screenshot=%s" % png, "--window-size=1600,900", "--virtual-time-budget=1500", "file://%s" % html_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(DEFAULT_OUT)); ap.add_argument("--themes", default="Modern,Boardroom")
    ap.add_argument("--write-abs", action="store_true"); ap.add_argument("--no-render", action="store_true"); ap.add_argument("--only")
    a = ap.parse_args()
    text = LAYOUTS.read_text(); spec = json.loads(text)
    base_grid = Grid(spec["slideSizes"]["Screen16x9"])
    sync_abs(text, base_grid, a.write_abs)
    spec = json.loads(LAYOUTS.read_text())
    themes, source = load_themes(); print("theme source:", source)
    out = Path(a.out); (out / "html").mkdir(parents=True, exist_ok=True); (out / "assets").mkdir(exist_ok=True)
    jobs = []
    for tn in a.themes.split(","):
        th = Theme(themes[tn], spec); slug = tn.lower()
        grid = base_grid.for_theme(themes[tn])
        S = samples(out)
        order = ["Title", "Section", "TitleAndContent", "TitleAndContent-table", "TitleAndContent-chart", "TwoContent", "TwoContent-textVisual", "Comparison", "Kpi", "Cards", "Process", "Quote", "ImageFull", "Closing", "CaseStudy", "Logos", "Coverage"]
        deck = Deck(spec, th, grid, "Quarterly business review", "Northwind Traders · Q3 2026 review")
        for i, key in enumerate(order):
            src = dict(S[key])
            if key == "Closing": idx, cnt = 9, 10
            elif key == "Title": idx, cnt = 0, 10
            else: idx, cnt = 3, 10
            lay = deck.select(src, idx, cnt)
            assert lay == key.split("-")[0], (key, lay)
            deck.section = "Performance" if key in ("TitleAndContent", "TitleAndContent-chart", "TwoContent", "Kpi", "Logos") else None
            if key == "Section": deck.section_no = 1
            for j, s in enumerate(deck.expand(dict(src, _layout=lay))):
                sl = render(deck, s, i + 1)
                jobs.append((write(out, "%s-%02d-%s" % (slug, i + 1, key), sl, key), out / ("%s-%02d-%s.png" % (slug, i + 1, key))))
        O = overflow_samples()
        for key, src in O.items():
            deck = Deck(spec, th, grid, "Overflow", "Overflow case · " + key); deck.section = None
            lay = deck.select(src, 3, 10)
            for j, s in enumerate(deck.expand(dict(src, _layout=lay))):
                sl = render(deck, s, j + 1)
                nm = "%s-overflow-%s-%d" % (slug, key, j + 1)
                jobs.append((write(out, nm, sl, nm), out / (nm + ".png")))
        rows = export_sample()
        src = export_deck(spec, rows, "Service review", ["Service", "Owner", "Status", "Tickets", "Availability"], "Service", "Tickets")
        deck = Deck(spec, th, grid, "Service review", "Service review · IT operations")
        built = deck.build(src)
        for j, s in enumerate(built):
            sl = render(deck, s, j + 1)
            nm = "%s-export-%d" % (slug, j + 1)
            jobs.append((write(out, nm, sl, nm), out / (nm + ".png")))
        print(tn, "export deck slides:", len(built))
    if a.no_render: return
    for h, p in jobs:
        if a.only and a.only not in p.name: continue
        shoot(h, p)
    print("rendered", len(jobs), "slides to", out)

if __name__ == "__main__":
    main()
