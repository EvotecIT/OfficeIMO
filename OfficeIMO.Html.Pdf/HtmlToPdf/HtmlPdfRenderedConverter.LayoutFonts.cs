using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>
    /// Freezes the existing library-selected Unicode fallback before authoritative layout.
    /// The provisional flow supplies visible text and coverage requirements before
    /// pagination; the caller reuses its resources and budget for one further layout.
    /// </summary>
    internal static bool TryPrepareLibraryFontMetrics(IEnumerable<HtmlRenderVisual> flow,
        IEnumerable<(string Text, OfficeFontInfo Font)> marginRequirements,
        OfficeFontFaceCollection fonts,
        HtmlToPdfOptions options, HtmlResourceSession resources, HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken) {
        if (options.LibrarySelectedDefaultFontFamily != null || options.FontFamily != null
            || !options.AllowSystemFontFallback
            || !options.ResourcePolicy.AllowSystemFontEmbedding || options.ResourcePolicy.AllowDocumentFontEmbedding
            || options.PdfOptions.HasEmbeddedStandardFontFamily(PdfCore.PdfStandardFont.Helvetica)
            || options.PdfOptions.HasNamedFontFamily(options.DefaultFontFamily)
            || options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) return false;

        var textRuns = EnumerateUsedText(flow).Concat(marginRequirements)
            .Where(text => !RetainedFacesCoverLayoutText(text.Text, text.Font, fonts, cancellationToken)).ToList();
        if (!textRuns.Any(text => RequiresUnicodeFont(text.Text))) return false;
        PdfCore.PdfEmbeddedFontFamily? family = SelectLibraryDefaultSystemFontFamily(cancellationToken);
        if (family == null) return false;
        family = CreateCoverageSafeFontFamily(family, textRuns);
        if (!TryRetainSelectedDefaultFontFamily(family, fonts, options, resources, diagnostics, cancellationToken)) return false;
        options.LibrarySelectedDefaultFontFamily = family;
        PdfCore.PdfCreateOptions measurement = options.PdfOptions.Clone();
        measurement.TextShapingLanguageOverride = options.TextShapingLanguage;
        measurement.SetTextShapingMode(options.TextShapingMode).SetTextShapingProvider(options.TextShapingProvider);
        measurement.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, options.LibrarySelectedDefaultFontFamily);
        ApplyLibraryDefaultFontEmbedding(options, measurement);
        options.FallbackTextMeasurement = CreateFallbackTextMeasurement(options, measurement);
        options.FallbackTextFaceMetrics = CreateFallbackTextFaceMetrics(options, measurement);
        return true;
    }

    // A retained family name alone does not prove coverage of a Unicode run.
    // Conversely, an exact non-static program can retain paint through outlines.
    private static bool RetainedFacesCoverLayoutText(string text, OfficeFontInfo font,
        OfficeFontFaceCollection fonts, CancellationToken cancellationToken) {
        foreach (OfficeFontFallbackRun run in fonts.PlanFallbackRuns(text, font.FamilyName, font.Style)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!fonts.TryResolveFaceForText(run.Text, run.FamilyRequest, font.Style, out _)) return false;
        }
        return true;
    }
}
