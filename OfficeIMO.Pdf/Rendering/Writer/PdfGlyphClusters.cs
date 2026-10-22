namespace OfficeIMO.Pdf;

/// <summary>Shared source-ownership boundaries for glyph emission and text distribution.</summary>
internal static class PdfGlyphClusters {
    /// <summary>Finds the end of one shaped source cluster, including overlapping GSUB source ownership.</summary>
    internal static int End(IReadOnlyList<PdfGlyphInfo> glyphs, int start, int limit) {
        int index = start;
        int logicalEnd = glyphs[index].LogicalClusterStart;
        do {
            PdfGlyphInfo glyph = glyphs[index++];
            // Multiple substitution followed by a ligature can leave different starts
            // whose source ranges overlap. Keep that entire ownership range together.
            logicalEnd = Math.Max(logicalEnd, Math.Max(glyph.LogicalClusterStart + 1,
                glyph.TextIndex + glyph.UnicodeText.Length));
        } while (index < limit && glyphs[index].LogicalClusterStart < logicalEnd);
        return index;
    }
}
