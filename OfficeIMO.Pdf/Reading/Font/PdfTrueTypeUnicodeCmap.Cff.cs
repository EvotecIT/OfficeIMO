using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfTrueTypeUnicodeCmap {
    private const uint OpenTypeCffFlavor = 0x4F54544FU;

    private static bool IsLoadable(byte[] program, bool cff) =>
        cff ? OfficeOpenTypeCffFont.TryLoad(program, null, out _) != null : OfficeTrueTypeFont.TryLoad(program) != null;

    /// <summary>
    /// Wraps a bare CFF program in a minimal OpenType ('OTTO') container whose Unicode cmap selects the glyph each
    /// PDF code paints, so CFF and converted Type 1 outlines use the shared OpenType CFF engine.
    /// </summary>
    /// <param name="font">The PDF font whose encoding or ToUnicode describes the painted text.</param>
    /// <param name="cff">The parsed bare CFF program.</param>
    /// <param name="glyphForCode">Simple fonts: the glyph a one-byte code paints. Ignored for CID fonts.</param>
    /// <param name="cidToGlyphMap">CID fonts: two bytes per CID naming its glyph; null for simple fonts.</param>
    /// <param name="advances">Glyph advances in glyph-space units of one em = <paramref name="unitsPerEm"/>.</param>
    /// <param name="unitsPerEm">Units per em implied by the font matrix.</param>
    /// <param name="originalProgramByteCount">Size of the embedded program retained by the glyph delegates.</param>
    internal static PdfDrawingFontProgram? TryCreateFromCff(PdfFontResource font, PdfCffFontProgram cff, Func<int, int>? glyphForCode,
        byte[]? cidToGlyphMap, double[] advances, int unitsPerEm, long originalProgramByteCount) {
        int glyphCount = cff.GlyphCount;
        Func<int, bool> isEmptyGlyph = cff.IsEmptyGlyph;
        SortedDictionary<int, int>? mappings;
        Func<int, int> codeToGlyph;
        if (cidToGlyphMap != null) {
            mappings = CreateCidMappings(font, cidToGlyphMap, glyphCount, isEmptyGlyph) ?? SeedMapping(glyphCount, isEmptyGlyph);
            codeToGlyph = cid => GlyphForCid(cidToGlyphMap, cid, glyphCount);
        } else if (glyphForCode != null) {
            mappings = CreateCodeMappings(font, glyphForCode, glyphCount, isEmptyGlyph);
            if (mappings.Count == 0) mappings = SeedMapping(glyphCount, isEmptyGlyph);
            codeToGlyph = code => code is < 0 or > 255 ? 0 : glyphForCode(code);
        } else {
            return null;
        }
        if (mappings == null || mappings.Count == 0 || mappings.Count > OfficeOpenTypeCmap.MaximumFormat12Groups) return null;
        byte[] program = BuildOpenTypeCff(cff, mappings, advances, unitsPerEm);
        return IsLoadable(program, cff: true)
            ? new PdfDrawingFontProgram(program, mappings, codeToGlyph, isEmptyGlyph, cidToGlyphMap, originalProgramByteCount)
            : null;
    }

    // Without ToUnicode no scalar names a glyph; one private-use entry registers the face and the
    // painted-glyph mapper aliases each painted CID to its glyph.
    private static SortedDictionary<int, int>? SeedMapping(int glyphCount, Func<int, bool> isEmptyGlyph) {
        for (int glyph = 1; glyph < glyphCount; glyph++) {
            if (!isEmptyGlyph(glyph)) return new SortedDictionary<int, int> { [0xE000] = glyph };
        }
        return glyphCount > 0 ? new SortedDictionary<int, int> { [0xE000] = 0 } : null;
    }

    private static SortedDictionary<int, int> CreateCodeMappings(PdfFontResource font, Func<int, int> glyphForCode, int glyphCount,
        Func<int, bool> isEmptyGlyph) {
        Func<byte[], int, string> decodeText = ResourceResolver.CreateBudgetedDecoder(font);
        var mappings = new SortedDictionary<int, int>();
        int fallbackGlyph = -1;
        for (int code = 0; code < 256; code++) {
            int glyph = glyphForCode(code);
            if (glyph <= 0 || glyph >= glyphCount) continue;
            if (fallbackGlyph < 0 && !isEmptyGlyph(glyph)) fallbackGlyph = glyph;
            string text;
            try {
                text = decodeText(new[] { (byte)code }, 8);
            } catch (Exception exception) when (exception is PdfReadLimitException or PdfUnsupportedTextMappingException) {
                continue;
            }
            if (!TryGetSingleScalar(text, out int scalar) || mappings.ContainsKey(scalar) ||
                IsWhiteSpaceScalar(scalar) && !isEmptyGlyph(glyph)) continue;
            mappings.Add(scalar, glyph);
        }
        if (mappings.Count == 0 && fallbackGlyph >= 0) mappings.Add(0xE000, fallbackGlyph);
        return mappings;
    }

    private static byte[] BuildOpenTypeCff(PdfCffFontProgram cff, SortedDictionary<int, int> mappings, double[] advances, int unitsPerEm) {
        int glyphCount = cff.GlyphCount;
        double scale = unitsPerEm * Math.Abs(cff.FontMatrix[0]);
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) scale = 1D;
        short xMin = ClampShort(cff.FontBBox[0] * scale), yMin = ClampShort(cff.FontBBox[1] * scale);
        short xMax = ClampShort(cff.FontBBox[2] * scale), yMax = ClampShort(cff.FontBBox[3] * scale);
        if (xMax <= xMin) { xMin = 0; xMax = (short)Math.Min(short.MaxValue, unitsPerEm); }
        if (yMax <= yMin) { yMin = (short)(-unitsPerEm / 5); yMax = (short)(unitsPerEm * 4 / 5); }
        short ascender = Math.Max(yMax, (short)1);
        short descender = Math.Min(yMin, (short)0);

        var hmtx = new byte[glyphCount * 4];
        int maximumAdvance = 0;
        for (int glyph = 0; glyph < glyphCount; glyph++) {
            double advance = glyph < advances.Length ? advances[glyph] : 0D;
            int value = double.IsNaN(advance) || double.IsInfinity(advance) ? 0 : (int)Math.Round(Math.Max(0D, Math.Min(ushort.MaxValue, advance)));
            maximumAdvance = Math.Max(maximumAdvance, value);
            hmtx[glyph * 4] = (byte)(value >> 8);
            hmtx[glyph * 4 + 1] = (byte)value;
        }

        var head = new byte[54];
        WriteUInt32(head, 0, 0x00010000U);
        WriteUInt32(head, 4, 0x00010000U);
        WriteUInt32(head, 12, 0x5F0F3CF5U);
        WriteUInt16(head, 16, 3);
        WriteUInt16(head, 18, (ushort)unitsPerEm);
        WriteUInt16(head, 36, (ushort)xMin);
        WriteUInt16(head, 38, (ushort)yMin);
        WriteUInt16(head, 40, (ushort)xMax);
        WriteUInt16(head, 42, (ushort)yMax);
        WriteUInt16(head, 46, 8);
        WriteUInt16(head, 48, 2);

        var hhea = new byte[36];
        WriteUInt32(hhea, 0, 0x00010000U);
        WriteUInt16(hhea, 4, (ushort)ascender);
        WriteUInt16(hhea, 6, (ushort)descender);
        WriteUInt16(hhea, 10, (ushort)maximumAdvance);
        WriteUInt16(hhea, 16, (ushort)xMax);
        WriteUInt16(hhea, 18, 1);
        WriteUInt16(hhea, 34, (ushort)glyphCount);

        var maxp = new byte[6];
        WriteUInt32(maxp, 0, 0x00005000U);
        WriteUInt16(maxp, 4, (ushort)glyphCount);

        var tables = new List<(string Tag, byte[] Body)> {
            ("CFF ", cff.Data),
            ("cmap", BuildUnicodeCmap(mappings)),
            ("head", head),
            ("hhea", hhea),
            ("hmtx", hmtx),
            ("maxp", maxp)
        };
        tables.Sort(static (left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        return Assemble(tables, OpenTypeCffFlavor);
    }

    private static short ClampShort(double value) =>
        double.IsNaN(value) ? (short)0 : (short)Math.Round(Math.Max(short.MinValue, Math.Min(short.MaxValue, value)));

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }
}
