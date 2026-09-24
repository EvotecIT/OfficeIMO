"""Create ligature-run.pdf: multi-glyph Latin runs that contain ligature glyphs.

Producers such as Microsoft Word and Ghostscript paint whole words in one string and map each
ligature glyph to its letters in ToUnicode ("ffi" for the ffi ligature). Drawing those letters would
replace the painted ligature glyph, so a renderer must draw each glyph by id. This fixture uses a
subset of the repository's Carlito (SIL Open Font License, see
Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/OFL-Carlito.txt) in a CIDFontType2 font with an
Identity CIDToGIDMap and paints "official", "fluffy" and "affinity" with their ligature glyphs.

Regenerate from the repository root with Python and fontTools 4.46.0 or newer:
    python OfficeIMO.Pdf.Tests/Pdf/Fixtures/ShapedText/create_ligature_run_fixture.py
"""
import io
import zlib
from pathlib import Path

from fontTools import subset
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parents[4]
SOURCE = ROOT / "Website" / "Apps" / "OfficeIMO.Web.Converter" / "Assets" / "Fonts" / "Carlito-Regular.ttf"
OUTPUT = Path(__file__).resolve().parent / "ligature-run.pdf"
LIGATURE_ONLY_OUTPUT = Path(__file__).resolve().parent / "ligature-only.pdf"
# Words as (glyph name, ToUnicode text) sequences; ligature glyphs map to several letters.
WORDS = [
    [("o", "o"), ("uniFB03", "ffi"), ("c", "c"), ("i", "i"), ("a", "a"), ("l", "l")],
    [("f", "f"), ("l", "l"), ("u", "u"), ("uniFB00", "ff"), ("y", "y")],
    [("a", "a"), ("uniFB03", "ffi"), ("n", "n"), ("i", "i"), ("t", "t"), ("y", "y")],
]

options = subset.Options()
options.notdef_outline = True
options.retain_gids = False
options.name_IDs = ["*"]
options.layout_features = ["liga"]
options.drop_tables += ["GPOS", "GDEF", "kern", "DSIG"]
font = TTFont(SOURCE)
subsetter = subset.Subsetter(options)
subsetter.populate(text="official fluffy affinity ")
subsetter.subset(font)
order = font.getGlyphOrder()
for word in WORDS:
    for name, _ in word:
        assert name in order, f"{name} is not in the subset: {order}"
gid = {name: order.index(name) for name in order}
font["cmap"].tables = [table for table in font["cmap"].tables if table.platformID == 3]
buffer = io.BytesIO()
font.save(buffer)
program = buffer.getvalue()

scale = 1000 / font["head"].unitsPerEm
def write_fixture(words, path, include_spaces):
    used = sorted({name for word in words for name, _ in word} | ({"space"} if include_spaces else set()),
                  key=lambda name: gid[name])
    to_unicode_entries = [(gid[name], text) for word in words for name, text in word]
    if include_spaces:
        to_unicode_entries.append((gid["space"], " "))
    unique = dict(sorted(to_unicode_entries))
    to_unicode = "\n".join([
        "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
        "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
        "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
        "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange", f"{len(unique)} beginbfchar",
        *[f"<{code:04X}> <{''.join(f'{ord(c):04X}' for c in text)}>" for code, text in unique.items()],
        "endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end", "",
    ]).encode("ascii")

    space = "<" + f"{gid['space']:04X}" + ">"
    line = " ".join("<" + "".join(f"{gid[name]:04X}" for name, _ in word) + "> Tj" +
                    (" " + space + " Tj" if include_spaces else "") for word in words)
    content = f"BT /F1 24 Tf 16 40 Td 0 0 0 rg {line} ET".encode("ascii")
    w_array = " ".join(f"{gid[name]} [{round(font['hmtx'][name][0] * scale)}]" for name in used)
    objects = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 90] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        5: b"<< /Type /Font /Subtype /Type0 /BaseFont /CARLIG+Carlito /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 8 0 R >>",
        6: (f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /CARLIG+Carlito "
            f"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /CIDToGIDMap /Identity "
            f"/DW 1000 /W [{w_array}] /FontDescriptor 7 0 R >>").encode("ascii"),
        7: b"<< /Type /FontDescriptor /FontName /CARLIG+Carlito /Flags 32 /FontBBox [-500 -300 1300 1000] "
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
    path.write_bytes(output.getvalue())
    print(f"Wrote {path} ({len(output.getvalue())} bytes)")


write_fixture(WORDS, OUTPUT, True)
write_fixture([[('uniFB03', 'ffi')]], LIGATURE_ONLY_OUTPUT, False)
