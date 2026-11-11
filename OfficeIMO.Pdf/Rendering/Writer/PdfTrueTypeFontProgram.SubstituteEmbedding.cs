namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    /// <summary>Returns the advance width in 1/1000 em without rounding, as validators compare it with PDF Widths.</summary>
    internal double GetExactGlyphWidth1000(int glyphId) {
        TableRecord hhea = GetTable(_tables, "hhea");
        TableRecord hmtx = GetTable(_tables, "hmtx");
        int numberOfHMetrics = ReadUInt16(_data, hhea.Offset + 34);
        int metric = Math.Min(Math.Max(glyphId, 0), numberOfHMetrics - 1);
        int offset = hmtx.Offset + metric * 4;
        EnsureRange(_data, offset, 4);
        return ReadUInt16(_data, offset) * 1000D / UnitsPerEm;
    }

    /// <summary>Builds a subset that keeps the original glyph numbering and cmap, for simple fonts addressed by glyph name.</summary>
    internal byte[] BuildGlyphRetainingSubset(IEnumerable<int> glyphIds) =>
        BuildSubsetFontFileUncached(glyphIds.Distinct().OrderBy(static glyph => glyph).ToArray(), compact: false, out _);
}
