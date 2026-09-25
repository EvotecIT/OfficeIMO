"""Print cairo-arabic-extended.pdf with Pango/Cairo, independently of OfficeIMO.

Pango shapes the text; Cairo embeds CIDFontType2 subsets whose ToUnicode maps each contextual
glyph back to its base letter. The page covers Persian and Urdu extended letters and one Arabic
word whose middle letters switch to a second font. Noto Sans Arabic comes from the repository;
the second font is DejaVu Sans from the system.

Run on Linux with python3-cairo, python3-gi-cairo, gir1.2-pango-1.0 and fonts-dejavu-core:
    python3 OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/generate_cairo_extended.py
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
NOTO_ARABIC = ROOT / "Website" / "Apps" / "OfficeIMO.Web.Converter" / "Assets" / "Fonts" / "NotoSansArabic-Regular.ttf"
LINES = [
    "پژوهش گزارش کیفیت چاپ",
    "یہ ٹیسٹ ڈیٹا ہے میں بڑی",
    "الم<span font_family=\"DejaVu Sans\">راجع</span>ة تبدأ",
]

# Make the repository font visible to fontconfig for this process only.
fonts = Path(tempfile.mkdtemp())
shutil.copy(NOTO_ARABIC, fonts / NOTO_ARABIC.name)
config = fonts / "fonts.conf"
config.write_text(f"""<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig><include ignore_missing="yes">/etc/fonts/fonts.conf</include><dir>{fonts}</dir><cachedir>{fonts}/cache</cachedir></fontconfig>
""")
os.environ["FONTCONFIG_FILE"] = str(config)

import cairo  # noqa: E402  (fontconfig must be configured first)
import gi  # noqa: E402
gi.require_version("Pango", "1.0")
gi.require_version("PangoCairo", "1.0")
from gi.repository import Pango, PangoCairo  # noqa: E402

output = HERE / "cairo-arabic-extended.pdf"
surface = cairo.PDFSurface(str(output), 300, 170)
surface.set_metadata(cairo.PDFMetadata.CREATE_DATE, "2026-09-24T00:00:00Z")
context = cairo.Context(surface)
y = 12
for line in LINES:
    layout = PangoCairo.create_layout(context)
    layout.set_font_description(Pango.FontDescription.from_string("Noto Sans Arabic 16"))
    layout.set_width(int(276 * Pango.SCALE))
    layout.set_alignment(Pango.Alignment.RIGHT)
    layout.set_auto_dir(True)
    layout.set_markup(line, -1)
    context.move_to(12, y)
    PangoCairo.show_layout(context, layout)
    y += 30
surface.finish()
shutil.rmtree(fonts)
print(subprocess.run(["pdffonts", str(output)], capture_output=True, text=True).stdout)
