"""Create space-coded-glyph.pdf: inked glyphs whose ToUnicode text is a space.

Colour-emoji and other layered output (for example Microsoft Word with Segoe UI Emoji) paints each
layer as its own glyph and maps every layer glyph to U+0020 in ToUnicode. A renderer must draw those
glyphs by glyph id while real spaces stay blank. This fixture reproduces that pattern with a subset of
OfficeIMO Baseline Sans (SIL Open Font License, see OfficeIMO.TestAssets/Fonts/OFL-Carlito.txt) in a
CIDFontType2 font with an Identity CIDToGIDMap: the space glyph and the inked "O" glyph both map to
U+0020, and the "O" is painted in red over each "K" like a colour layer.

Regenerate from the repository root with Python and fontTools 4.46.0 or newer:
    python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_space_coded_glyph_fixture.py
"""
import io
import zlib
from pathlib import Path

from fontTools import subset
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parents[4]
SOURCE = ROOT / "OfficeIMO.TestAssets" / "Fonts" / "OfficeIMOBaselineSans-Regular.ttf"
OUTPUT = Path(__file__).resolve().parent / "space-coded-glyph.pdf"

options = subset.Options()
options.notdef_outline = True
options.retain_gids = False
options.name_IDs = ["*"]
options.drop_tables += ["GSUB", "GPOS", "GDEF", "kern", "DSIG"]
font = TTFont(SOURCE)
subsetter = subset.Subsetter(options)
subsetter.populate(text=" KO")
subsetter.subset(font)
cmap = font.getBestCmap()
gid = {character: font.getGlyphID(cmap[ord(character)]) for character in " KO"}
buffer = io.BytesIO()
font.save(buffer)
program = buffer.getvalue()

scale = 1000 / font["head"].unitsPerEm
width = {character: round(font["hmtx"][cmap[ord(character)]][0] * scale) for character in " KO"}
to_unicode = "\n".join([
    "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
    "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
    "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
    "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange", "3 beginbfchar",
    f"<{gid[' ']:04X}> <0020>", f"<{gid['O']:04X}> <0020>", f"<{gid['K']:04X}> <004B>",
    "endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end", "",
]).encode("ascii")


def glyphs(text):
    return "<" + "".join(f"{gid[c]:04X}" for c in text) + ">"


# Each "K" is drawn in black, then the space-coded "O" layer in red at the same origin.
content = "\n".join([
    "BT /F1 36 Tf 20 30 Td 0 0 0 rg " + glyphs("K") + " Tj ET",
    "BT /F1 36 Tf 20 30 Td 0.85 0.1 0.1 rg " + glyphs("O") + " Tj ET",
    "BT /F1 36 Tf 80 30 Td 0 0 0 rg " + glyphs("K K") + " Tj ET",
    "BT /F1 36 Tf 80 30 Td 0.85 0.1 0.1 rg " + glyphs("O") + " Tj ET",
]).encode("ascii")
w_array = " ".join(f"{gid[c]} [{width[c]}]" for c in " KO")
objects = {
    1: b"<< /Type /Catalog /Pages 2 0 R >>",
    2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 80] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    5: b"<< /Type /Font /Subtype /Type0 /BaseFont /OFIMOS+OfficeIMOBaselineSans /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 8 0 R >>",
    6: (f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /OFIMOS+OfficeIMOBaselineSans "
        f"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /CIDToGIDMap /Identity "
        f"/DW 1000 /W [{w_array}] /FontDescriptor 7 0 R >>").encode("ascii"),
    7: b"<< /Type /FontDescriptor /FontName /OFIMOS+OfficeIMOBaselineSans /Flags 32 /FontBBox [-500 -300 1300 1000] "
       b"/ItalicAngle 0 /Ascent 950 /Descent -250 /CapHeight 630 /StemV 80 /FontFile2 9 0 R >>",
}
streams = {4: content, 8: to_unicode, 9: zlib.compress(program)}

output = io.BytesIO()
output.write(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
offsets = {}
for number in range(1, 10):
    offsets[number] = output.tell()
    output.write(f"{number} 0 obj\n".encode("ascii"))
    if number in streams:
        data = streams[number]
        extra = f" /Filter /FlateDecode /Length1 {len(program)}" if number == 9 else ""
        output.write(f"<< /Length {len(data)}{extra} >>\nstream\n".encode("ascii") + data + b"\nendstream")
    else:
        output.write(objects[number])
    output.write(b"\nendobj\n")
xref = output.tell()
output.write(b"xref\n0 10\n0000000000 65535 f \n")
for number in range(1, 10):
    output.write(f"{offsets[number]:010d} 00000 n \n".encode("ascii"))
output.write(f"trailer\n<< /Size 10 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("ascii"))
OUTPUT.write_bytes(output.getvalue())
print(f"Wrote {OUTPUT} ({len(output.getvalue())} bytes)")
