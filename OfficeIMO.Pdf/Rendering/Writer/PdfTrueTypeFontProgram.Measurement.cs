namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
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
}
