"""Create symbolic-truetype-cmap.pdf: a simple TrueType font whose only cmap is (3,0).

The embedded program is a subset of OfficeIMO Baseline Sans (SIL Open Font License, see
OfficeIMO.TestAssets/Fonts/OFL-Carlito.txt). Its Unicode cmap is replaced by a Windows symbol
(3,0) format 4 subtable keyed by 0xF000 + PDF character code, as producers of symbolic subsets
write it. The PDF has symbolic font flags, no /Encoding, and a ToUnicode map for extraction.

Regenerate from the repository root with Python and fontTools 4.46.0 or newer:
    python OfficeIMO.Pdf.Tests/Pdf/Fixtures/Fonts/create_symbolic_truetype_fixture.py
"""
import io
import zlib
from pathlib import Path

from fontTools import subset
from fontTools.ttLib import TTFont
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

ROOT = Path(__file__).resolve().parents[4]
SOURCE = ROOT / "OfficeIMO.TestAssets" / "Fonts" / "OfficeIMOBaselineSans-Regular.ttf"
OUTPUT = Path(__file__).resolve().parent / "symbolic-truetype-cmap.pdf"
TEXT = "Symbolic cmap glyphs"

characters = sorted(set(TEXT))
codes = {character: 0x20 + index for index, character in enumerate(characters)}

options = subset.Options()
options.notdef_outline = True
options.name_IDs = ["*"]
options.drop_tables += ["GSUB", "GPOS", "GDEF", "kern", "DSIG"]
font = TTFont(SOURCE)
subsetter = subset.Subsetter(options)
subsetter.populate(text=TEXT)
subsetter.subset(font)

unicode_cmap = font.getBestCmap()
symbol = CmapSubtable.newSubtable(4)
symbol.platformID, symbol.platEncID, symbol.language = 3, 0, 0
symbol.cmap = {0xF000 + codes[character]: unicode_cmap[ord(character)] for character in characters}
font["cmap"].tables = [symbol]
buffer = io.BytesIO()
font.save(buffer)
program = buffer.getvalue()

scale = 1000 / font["head"].unitsPerEm
widths = [round(font["hmtx"][unicode_cmap[ord(character)]][0] * scale) for character in characters]
first, last = codes[characters[0]], codes[characters[-1]]
encoded = bytes(codes[character] for character in TEXT)
to_unicode = "\n".join([
    "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
    "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
    "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
    "1 begincodespacerange", "<00> <FF>", "endcodespacerange",
    f"{len(characters)} beginbfchar",
    *[f"<{codes[c]:02X}> <{ord(c):04X}>" for c in characters],
    "endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end", "",
]).encode("ascii")
content = b"BT /F1 24 Tf 20 40 Td <" + encoded.hex().upper().encode("ascii") + b"> Tj ET"

objects = [
    b"<< /Type /Catalog /Pages 2 0 R >>",
    b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    None,
    (f"<< /Type /Font /Subtype /TrueType /BaseFont /OFIMOS+OfficeIMOBaselineSans /FirstChar {first} /LastChar {last} "
     f"/Widths [{' '.join(str(width) for width in widths)}] /FontDescriptor 6 0 R /ToUnicode 7 0 R >>").encode("ascii"),
    b"<< /Type /FontDescriptor /FontName /OFIMOS+OfficeIMOBaselineSans /Flags 4 /FontBBox [-500 -300 1300 1000] "
    b"/ItalicAngle 0 /Ascent 950 /Descent -250 /CapHeight 630 /StemV 80 /FontFile2 8 0 R >>",
    None,
    None,
]
streams = {4: content, 7: to_unicode, 8: zlib.compress(program)}

output = io.BytesIO()
output.write(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
offsets = []
for number in range(1, len(objects) + 1):
    offsets.append(output.tell())
    output.write(f"{number} 0 obj\n".encode("ascii"))
    if number in streams:
        data = streams[number]
        extra = f" /Filter /FlateDecode /Length1 {len(program)}" if number == 8 else ""
        output.write(f"<< /Length {len(data)}{extra} >>\nstream\n".encode("ascii") + data + b"\nendstream")
    else:
        output.write(objects[number - 1])
    output.write(b"\nendobj\n")
xref = output.tell()
output.write(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
for offset in offsets:
    output.write(f"{offset:010d} 00000 n \n".encode("ascii"))
output.write(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("ascii"))
OUTPUT.write_bytes(output.getvalue())
print(f"Wrote {OUTPUT} ({len(output.getvalue())} bytes)")
