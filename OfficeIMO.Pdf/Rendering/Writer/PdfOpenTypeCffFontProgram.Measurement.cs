namespace OfficeIMO.Pdf;

internal sealed partial class PdfOpenTypeCffFontProgram {
    private int MeasureDefaultLatinAdvanceWidth1000(string text, PdfTextShapingOptions options) {
        if (!PdfShortTextCache<PdfMeasuredText>.IsEligible(text, options))
            return PdfExternalTextShaper.MeasureDefaultLatinAdvanceWidth1000(text, this, options);
        if (_shortMeasurements.TryGet(text, options, out PdfMeasuredText cached)) {
            cached.ReplayUsage(RecordGlyphUsage);
            return cached.Advance;
        }
        var usage = new List<(int, string)>(text.Length);
        int advance = PdfExternalTextShaper.MeasureDefaultLatinAdvanceWidth1000(text, this, options,
            (glyphId, unicode) => usage.Add((glyphId, unicode)));
        _shortMeasurements.Add(text, options, new PdfMeasuredText(advance, usage), usage.Count);
        return advance;
    }

    // Mirrors the CFF scalar fallback, including its missing-glyph exception and usage
    // recording, without materializing a run solely to sum nominal advances.
    internal int MeasureScalarAdvanceWidth1000(string text, PdfTextShapingOptions options, Action<int, string>? observeUsage = null) {
        int total = 0;
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (!TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                if (options.ThrowOnMissingGlyph) throw CreateUnsupportedGlyphException(text, scalarStart, scalar);
                continue;
            }
            if (options.RecordGlyphUsage) RecordGlyphUsage(glyphId, scalar);
            observeUsage?.Invoke(glyphId, char.ConvertFromUtf32(scalar));
            total = checked(total + GetGlyphWidth1000(glyphId));
        }
        return total;
    }
}
