namespace OfficeIMO.Pdf;

internal static partial class PdfTrueTypeUnicodeCmap {
    private const int MaximumReverseCmapCodes = 0x110000;

    /// <summary>
    /// Reads the program's Unicode cmap backwards: the Unicode scalar of each glyph, or null when the program has
    /// no usable Unicode subtable. A glyph reached from several scalars keeps the lowest one outside the private use
    /// areas, so a space shared with no-break space reads as a space and a CJK compatibility alias as its unified form.
    /// </summary>
    internal static Dictionary<int, int>? TryReadGlyphScalars(byte[] program) {
        if (!TryReadTables(program, out List<(string Tag, int Offset, int Length)> tables)) return null;
        int cmapIndex = tables.FindIndex(static table => string.Equals(table.Tag, "cmap", StringComparison.Ordinal));
        if (cmapIndex < 0 ||
            !TryReadCodeSubtables(program, tables[cmapIndex].Offset, tables[cmapIndex].Length, out _, out _, out int? unicode) ||
            unicode is not int table) return null;
        var glyphs = new Dictionary<int, int>();
        switch (ReadUInt16(program, table)) {
            case 4:
                ReadFormat4(program, table, glyphs);
                break;
            case 12:
                ReadFormat12(program, table, glyphs);
                break;
        }
        return glyphs.Count == 0 ? null : glyphs;
    }

    private static void ReadFormat4(byte[] data, int table, Dictionary<int, int> glyphs) {
        int length = ReadUInt16(data, table + 2);
        if (length < 16) return;
        int segments = ReadUInt16(data, table + 6) / 2;
        int ends = table + 14;
        int starts = ends + segments * 2 + 2;
        int deltas = starts + segments * 2;
        int rangeOffsets = deltas + segments * 2;
        if (rangeOffsets + segments * 2 > table + length) return;
        for (int segment = 0; segment < segments; segment++) {
            int end = ReadUInt16(data, ends + segment * 2);
            int start = ReadUInt16(data, starts + segment * 2);
            int delta = ReadUInt16(data, deltas + segment * 2);
            int rangeOffsetPosition = rangeOffsets + segment * 2;
            int rangeOffset = ReadUInt16(data, rangeOffsetPosition);
            for (int code = start; code <= end && code < 0xFFFF; code++) {
                int glyph;
                if (rangeOffset == 0) {
                    glyph = (code + delta) & 0xFFFF;
                } else {
                    int glyphPosition = rangeOffsetPosition + rangeOffset + (code - start) * 2;
                    if (glyphPosition > table + length - 2) break;
                    glyph = ReadUInt16(data, glyphPosition);
                    if (glyph != 0) glyph = (glyph + delta) & 0xFFFF;
                }
                AddGlyphScalar(glyphs, glyph, code);
            }
        }
    }

    private static void ReadFormat12(byte[] data, int table, Dictionary<int, int> glyphs) {
        long groups = ReadUInt32(data, table + 12);
        long visited = 0;
        for (long group = 0; group < groups; group++) {
            int record = table + 16 + (int)group * 12;
            long start = ReadUInt32(data, record);
            long end = ReadUInt32(data, record + 4);
            long startGlyph = ReadUInt32(data, record + 8);
            for (long code = start; code <= end && code <= 0x10FFFF; code++) {
                long glyph = startGlyph + (code - start);
                if (glyph > ushort.MaxValue || ++visited > MaximumReverseCmapCodes) break;
                AddGlyphScalar(glyphs, (int)glyph, (int)code);
            }
            if (visited > MaximumReverseCmapCodes) return;
        }
    }

    private static void AddGlyphScalar(Dictionary<int, int> glyphs, int glyph, int scalar) {
        if (glyph == 0 || !IsTextScalar(scalar)) return;
        if (!glyphs.TryGetValue(glyph, out int existing) || IsPrivateUse(existing) && !IsPrivateUse(scalar)) glyphs[glyph] = scalar;
    }

    private static bool IsTextScalar(int scalar) =>
        scalar >= 0x20 && scalar <= 0x10FFFF && scalar != 0xFFFD && !(scalar >= 0x7F && scalar < 0xA0) &&
        !(scalar >= 0xD800 && scalar <= 0xDFFF) && (scalar & 0xFFFE) != 0xFFFE && !(scalar >= 0xFDD0 && scalar <= 0xFDEF);

    private static bool IsPrivateUse(int scalar) =>
        scalar is >= 0xE000 and <= 0xF8FF or >= 0xF0000;
}
