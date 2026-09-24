"""Regenerate the painted-glyph rendering references.

Poppler (an independent PDF renderer) rasterizes page 1 of each case at 144 dpi. The raster is
thresholded at luminance 128 and stored as a 1-bit PNG, the form PdfPaintedGlyphRenderingTests
compares against. Pass --chrome to reprint the Chrome fixtures from their HTML sources first.

chrome-arabic-extended.html paints part of one word in a second font. Chrome shares one subset
between faces with identical bytes, so the script writes a renamed copy of Noto Sans Arabic
(fixture-arabic-second.ttf) for printing and removes it afterwards.

Requires Poppler's pdftoppm on PATH, Python Pillow, and fontTools for --chrome. Run from the
repository root:
    python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_references.py [--chrome <chrome.exe>]
"""
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
NOTO_ARABIC = ROOT / "Website" / "Apps" / "OfficeIMO.Web.Converter" / "Assets" / "Fonts" / "NotoSansArabic-Regular.ttf"
CHROME_FIXTURES = ("chrome-arabic", "chrome-arabic-extended", "chrome-devanagari")
CASES = {
    "chrome-arabic": HERE / "chrome-arabic.pdf",
    "chrome-arabic-extended": HERE / "chrome-arabic-extended.pdf",
    "cairo-arabic-extended": HERE / "cairo-arabic-extended.pdf",
    "chrome-devanagari": HERE / "chrome-devanagari.pdf",
    "space-coded-glyph": HERE / "space-coded-glyph.pdf",
    "ligature-run": HERE / "ligature-run.pdf",
    "ghostscript-ligature-run": HERE / "ghostscript-ligature-run.pdf",
    "cairo-rtl-0": ROOT / "OfficeIMO.TestAssets" / "MultilingualLayout" / "rtl-0-native.pdf",
    "cairo-rtl-90": ROOT / "OfficeIMO.TestAssets" / "MultilingualLayout" / "rtl-90-native.pdf",
    "cairo-latin-0": ROOT / "OfficeIMO.TestAssets" / "MultilingualLayout" / "latin-0-native.pdf",
    "word-mac-report": ROOT / "OfficeIMO.Pdf.Tests" / "Pdf" / "ReferenceBaselines" / "microsoft-word-16.109-native-word-report.pdf",
    "word-windows-summary": ROOT / "OfficeIMO.Pdf.Tests" / "Pdf" / "ReferenceBaselines" / "microsoft-word-windows-word-business-delivery-summary.pdf",
}


def write_second_font(path):
    from fontTools.ttLib import TTFont

    font = TTFont(NOTO_ARABIC)
    for record in font["name"].names:
        if record.nameID in (1, 3, 4, 6, 16):
            record.string = str(record.toUnicode()).replace("Noto Sans Arabic", "Fixture Arabic Second").replace(
                "NotoSansArabic", "FixtureArabicSecond")
    font.save(path)


if "--chrome" in sys.argv:
    chrome = sys.argv[sys.argv.index("--chrome") + 1]
    second = HERE / "fixture-arabic-second.ttf"
    write_second_font(second)
    try:
        for name in CHROME_FIXTURES:
            subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                            "--allow-file-access-from-files", f"--print-to-pdf={HERE / (name + '.pdf')}",
                            (HERE / (name + ".html")).as_uri()], check=True)
    finally:
        second.unlink(missing_ok=True)

with tempfile.TemporaryDirectory() as temporary:
    for name, pdf in CASES.items():
        prefix = Path(temporary) / name
        subprocess.run(["pdftoppm", "-r", "144", "-f", "1", "-l", "1", "-singlefile", "-gray", "-png", str(pdf), str(prefix)], check=True)
        with Image.open(str(prefix) + ".png") as raster:
            raster.convert("L").point(lambda value: 0 if value < 128 else 255).convert("1").save(HERE / f"poppler-{name}.png", optimize=True)
        print(f"poppler-{name}.png")
