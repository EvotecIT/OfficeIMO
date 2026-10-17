using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>
    /// Freezes the existing library-selected Unicode fallback before authoritative layout.
    /// The provisional flow supplies visible text and coverage requirements before
    /// pagination; the caller reuses its resources and budget for one further layout.
    /// </summary>
    internal static bool TryPrepareLibraryFontMetrics(IEnumerable<HtmlRenderVisual> flow,
        OfficeFontFaceCollection fonts,
        HtmlToPdfOptions options, CancellationToken cancellationToken) {
        if (options.LibrarySelectedDefaultFontFamily != null || options.FontFamily != null
            || !options.ResourcePolicy.AllowSystemFontEmbedding || options.ResourcePolicy.AllowDocumentFontEmbedding
            || options.PdfOptions.HasEmbeddedStandardFontFamily(PdfCore.PdfStandardFont.Helvetica)
            || options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) return false;

        var webFamilies = new HashSet<string>(fonts.Faces.Where(face => face.CanEmbedAsStaticPdfFont)
            .Select(face => face.ResourceFamilyName), StringComparer.OrdinalIgnoreCase);
        var textRuns = EnumerateUsedText(flow)
            .Where(text => !EnumerateFamilies(text.Font.FamilyName).Any(webFamilies.Contains)).ToList();
        if (!textRuns.Any(text => RequiresUnicodeFont(text.Text))) return false;
        PdfCore.PdfEmbeddedFontFamily? family = SelectLibraryDefaultSystemFontFamily(cancellationToken);
        if (family == null) return false;
        options.LibrarySelectedDefaultFontFamily = CreateCoverageSafeFontFamily(family, textRuns);
        PdfCore.PdfOptions measurement = options.PdfOptions.Clone();
        measurement.TextShapingLanguageOverride = options.TextShapingLanguage;
        measurement.SetTextShapingMode(options.TextShapingMode).SetTextShapingProvider(options.TextShapingProvider);
        measurement.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, options.LibrarySelectedDefaultFontFamily);
        options.FallbackTextMeasurement = CreateFallbackTextMeasurement(options, measurement);
        options.FallbackTextFaceMetrics = CreateFallbackTextFaceMetrics(options, measurement);
        return true;
    }
}
