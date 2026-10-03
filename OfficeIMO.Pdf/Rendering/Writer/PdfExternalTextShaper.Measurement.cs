using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfExternalTextShaper {
    internal static int MeasureDefaultLatinAdvanceWidth1000(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, Action<int, string>? observeUsage = null) {
        var request = CreateDefaultLatinMeasurementRequest(text, font.FontName, font.FontDataForInspection, false, font.UnitsPerEm, options.Language);
        if (OfficeManagedTextShapingProvider.Instance.TryMeasureDefaultLatinText(request, (glyphId, unicodeText) => {
            ValidateMeasuredGlyph(glyphId, font.GlyphCount);
            int width = font.GetGlyphWidth1000(glyphId);
            if (options.RecordGlyphUsage) font.RecordGlyphUsage(glyphId, unicodeText);
            observeUsage?.Invoke(glyphId, unicodeText);
            return width;
        }, out int advance)) return advance;
        return PdfUnicodeScalarTextShaper.MeasureAdvanceWidth1000(text, font, options, observeUsage);
    }

    internal static int MeasureDefaultLatinAdvanceWidth1000(string text, PdfOpenTypeCffFontProgram font, PdfTextShapingOptions options, Action<int, string>? observeUsage = null) {
        var request = CreateDefaultLatinMeasurementRequest(text, font.FontName, font.FontDataForInspection, true, font.UnitsPerEm, options.Language);
        if (OfficeManagedTextShapingProvider.Instance.TryMeasureDefaultLatinText(request, (glyphId, unicodeText) => {
            ValidateMeasuredGlyph(glyphId, font.GlyphCount);
            int width = font.GetGlyphWidth1000(glyphId);
            if (options.RecordGlyphUsage) font.RecordGlyphUsage(glyphId, unicodeText);
            observeUsage?.Invoke(glyphId, unicodeText);
            return width;
        }, out int advance)) return advance;
        return font.MeasureScalarAdvanceWidth1000(text, options, observeUsage);
    }

    private static OfficeTextShapingRequest CreateDefaultLatinMeasurementRequest(string text, string fontName, byte[] fontData, bool isCff, int unitsPerEm, string? language) =>
        new OfficeTextShapingRequest(text, fontName, fontData, isCff, unitsPerEm,
            OfficeTextDirection.Auto, language, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: false, applyDefaultLatinLigatures: true);

    private static void ValidateMeasuredGlyph(int glyphId, int glyphCount) {
        if (glyphId <= 0 || glyphId >= glyphCount) throw new ArgumentException(
            "PDF text shaping provider returned glyph id " + glyphId.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            ", which is outside the embedded font glyph range.", nameof(glyphId));
    }
}
