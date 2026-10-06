namespace OfficeIMO.Pdf;

public sealed partial class PdfTextSpan {
    /// <summary>
    /// Projects the reader's glyph codes, text ranges and baseline origins into immutable evidence.
    /// Returns false with an empty list when glyph evidence is unavailable or the logical text
    /// cannot be mapped to painted glyphs, including ambiguous ActualText replacements.
    /// The origins describe the baseline; they are not exact glyph outlines or clipping bounds.
    /// </summary>
    public bool TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs) {
        glyphs = Array.Empty<PdfTextGlyph>();
        if (HasActualText || GlyphCharacterLengths is null || GlyphBytes is null || GlyphPaintedAdvances is null ||
            GlyphCharacterLengths.Count != GlyphBytes.Count || GlyphBytes.Count != GlyphPaintedAdvances.Count ||
            GlyphCharacterLengths.Count == 0 || GlyphCharacterLengths.Any(static length => length <= 0) ||
            GlyphCharacterLengths.Sum(static length => (long)length) != Text.Length ||
            !PdfTextAdvanceProjection.TryGetResolvedBoundaries(this, out double[] boundaries)) return false;

        var result = new List<PdfTextGlyph>(GlyphBytes.Count);
        double angle = RotationDegrees * Math.PI / 180D;
        double ux = Math.Cos(angle), uy = Math.Sin(angle);
        int start = 0;
        for (int index = 0; index < GlyphBytes.Count; index++) {
            int length = GlyphCharacterLengths[index];
            double offset = boundaries[start];
            result.Add(new PdfTextGlyph(Text.Substring(start, length), start,
                X + ux * offset, Y + uy * offset, boundaries[start + length] - offset,
                GlyphPaintedAdvances[index], GlyphBytes[index]));
            start += length;
        }
        glyphs = result.AsReadOnly();
        return true;
    }
}
