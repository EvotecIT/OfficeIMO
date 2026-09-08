"""Controlled layout corpus rendered by Pango/Cairo, independently of OfficeIMO.

Run on Linux with python3-cairo, python3-gi-cairo and gir1.2-pango-1.0.
The authoring records below also define the expected logical reading order.
"""
import hashlib
import json
import math
from pathlib import Path

import cairo
import gi
gi.require_version("Pango", "1.0")
gi.require_version("PangoCairo", "1.0")
from gi.repository import Pango, PangoCairo

ROOT = Path(__file__).resolve().parent
WIDTH, HEIGHT, SCALE = 595, 842, 300 / 72
CASES = [
    {
        "id": "latin", "languages": "eng+pol+deu", "rtl": False,
        "title": "Raport jakości dokumentu",
        "first": ["Pierwsza kolumna zawiera opis.", "Zażółć gęślą jaźń.",
                  "Prüfung der Qualität folgt.", "Die nächste Zeile bleibt links."],
        "second": ["Second column starts here.", "Version Office IMO 12.5 is ready.",
                   "Dane końcowe są poprawne.", "Last paragraph closes the report."],
        "caption": "Rysunek 1: wynik pomiaru",
        "table": [["Produkt", "Liczba"], ["Alfa", "24"], ["Beta", "12"], ["Gamma", "36"]],
    },
    {
        "id": "rtl", "languages": "heb+ara+eng", "rtl": True,
        "title": "דוח בדיקת מסמכים",
        "first": ["העמודה הראשונה מתחילה כאן.", "גרסה Office IMO 12.5 מוכנה.",
                  "תוצאת הבדיקה נשמרת במסמך.", "השורה האחרונה נשארת מימין."],
        "second": ["تقرير المراجعة يبدأ هنا.", "الإصدار Office IMO 12.5 جاهز.",
                   "نتيجة الاختبار محفوظة.", "الفقرة الأخيرة تغلق التقرير."],
        "caption": "איור 1: תוצאת המדידה",
        "table": [["פריט", "כמות"], ["אלפא", "24"], ["בטא", "12"], ["גמא", "36"]],
    },
]


def text(ctx, value, x, y, width=240, size=11, bold=False, rtl=False):
    layout = PangoCairo.create_layout(ctx)
    layout.get_context().set_base_dir(Pango.Direction.RTL if rtl else Pango.Direction.LTR)
    layout.set_auto_dir(False)
    layout.set_font_description(Pango.FontDescription(f"DejaVu Sans {'Bold ' if bold else ''}{size}"))
    layout.set_width(int(width * Pango.SCALE))
    layout.set_alignment(Pango.Alignment.RIGHT if rtl else Pango.Alignment.LEFT)
    layout.set_text(value, -1)
    if layout.get_line_count() != 1:
        raise ValueError(f"Unexpected wrapping: {value}")
    ctx.move_to(x, y)
    PangoCairo.show_layout(ctx, layout)


def figure():
    surface = cairo.ImageSurface(cairo.FORMAT_RGB24, 256, 96)
    ctx = cairo.Context(surface)
    ctx.set_source_rgb(.9, .95, 1)
    ctx.paint()
    for index, value in enumerate([24, 12, 36]):
        ctx.set_source_rgb(.1, .3 + index * .1, .65)
        ctx.rectangle(24 + index * 72, 80 - value * 1.7, 42, value * 1.7)
        ctx.fill()
    return surface


def draw(ctx, case):
    ctx.set_source_rgb(1, 1, 1)
    ctx.paint()
    ctx.set_source_rgb(0, 0, 0)
    rtl = case["rtl"]
    text(ctx, case["title"], 45, 50, 505, 17, True, rtl)
    first_x, second_x = (315, 45) if rtl else (45, 315)
    # Deliberately interleave column paint order. Expected reading order is column-major.
    for index in range(4):
        y = 120 + index * 42
        text(ctx, case["first"][index], first_x, y, 240, 10, rtl=rtl)
        text(ctx, case["second"][index], second_x, y, 240, 10, rtl=rtl)
    table = case["table"]
    for row, cells in enumerate(table):
        for column, value in enumerate(cells):
            x = (340 if column == 0 else 180) if rtl else (180 if column == 0 else 340)
            text(ctx, value, x, 355 + row * 28, 90, 11, row == 0, rtl)
    image = figure()
    ctx.save()
    ctx.translate(170, 530)
    ctx.set_source_surface(image)
    ctx.paint()
    ctx.restore()
    text(ctx, case["caption"], 170, 632, 256, 10, rtl=rtl)


def render(case, turn, raster):
    width, height = (HEIGHT, WIDTH) if turn % 2 else (WIDTH, HEIGHT)
    stem = f"{case['id']}-{turn * 90}"
    if raster:
        surface = cairo.ImageSurface(cairo.FORMAT_RGB24, math.ceil(width * SCALE), math.ceil(height * SCALE))
        ctx = cairo.Context(surface)
        ctx.scale(SCALE, SCALE)
    else:
        surface = cairo.PDFSurface(str(ROOT / f"{stem}-native.pdf"), width, height)
        surface.set_metadata(cairo.PDF_METADATA_CREATE_DATE, "2026-09-08T00:00:00Z")
        surface.set_metadata(cairo.PDF_METADATA_MOD_DATE, "2026-09-08T00:00:00Z")
        ctx = cairo.Context(surface)
    if turn == 1:
        ctx.translate(HEIGHT, 0)
    elif turn == 2:
        ctx.translate(WIDTH, HEIGHT)
    elif turn == 3:
        ctx.translate(0, WIDTH)
    ctx.rotate(turn * math.pi / 2)
    draw(ctx, case)
    if raster:
        surface.write_to_png(str(ROOT / f"{stem}-scan.png"))
        pdf = cairo.PDFSurface(str(ROOT / f"{stem}-scan.pdf"), width, height)
        # Searchable stamping currently supports this classic-xref container. Keep native
        # reading fixtures at Cairo's default version; do not rewrite measured input in the runner.
        pdf.restrict_to_version(cairo.PDF_VERSION_1_4)
        pdf.set_metadata(cairo.PDF_METADATA_CREATE_DATE, "2026-09-08T00:00:00Z")
        pdf.set_metadata(cairo.PDF_METADATA_MOD_DATE, "2026-09-08T00:00:00Z")
        pdfctx = cairo.Context(pdf)
        pdfctx.scale(1 / SCALE, 1 / SCALE)
        pdfctx.set_source_surface(surface)
        pdfctx.paint()
        pdf.finish()
    surface.finish()


manifest = {"producer": "Pango/Cairo", "cairo": cairo.cairo_version_string(),
            "pango": Pango.version_string(), "dpi": 300, "cases": []}
for case in CASES:
    for turn in range(4):
        render(case, turn, False)
        render(case, turn, True)
        stem = f"{case['id']}-{turn * 90}"
        order = [case["title"], *case["first"], *case["second"],
                 *[" ".join(row) for row in case["table"]], case["caption"]]
        manifest["cases"].append({"id": stem, "languages": case["languages"],
            "clockwiseDegrees": turn * 90, "rightToLeft": case["rtl"],
            "readingOrder": order, "table": case["table"], "caption": case["caption"],
            "files": {suffix: hashlib.sha256((ROOT / f"{stem}-{suffix}").read_bytes()).hexdigest()
                      for suffix in ["native.pdf", "scan.pdf", "scan.png"]}})
(ROOT / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
