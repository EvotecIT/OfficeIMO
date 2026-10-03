namespace OfficeIMO.Pdf;

internal sealed partial class PdfOpenTypeCffFontProgram {
    // Mirrors the CFF scalar fallback, including its missing-glyph exception and usage
    // recording, without materializing a run solely to sum nominal advances.
    internal int MeasureScalarAdvanceWidth1000(string text, PdfTextShapingOptions options) {
        int total = 0;
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (!TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                if (options.ThrowOnMissingGlyph) throw CreateUnsupportedGlyphException(text, scalarStart, scalar);
                continue;
            }
            if (options.RecordGlyphUsage) RecordGlyphUsage(glyphId, scalar);
            total = checked(total + GetGlyphWidth1000(glyphId));
        }
        return total;
    }
}
