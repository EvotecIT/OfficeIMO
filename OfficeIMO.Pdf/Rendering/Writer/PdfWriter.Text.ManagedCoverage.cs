namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static PdfTextShapingOptions? CreateSelectedManagedCoverageOptions(PdfCreateOptions options,
        string fontName, PdfTextRun run, bool preferSelected) => preferSelected
            ? PdfTextDiagnostics.ForEncodingDiagnostics(options, fontName,
                featureSettings: run.FeatureSettings, direction: run.TextDirection,
                preservesPaintedGlyphs: run.PreservesPaintedGlyphs)
            : null;

    private static bool TryGetSelectedManagedCoverage(string text, PdfTrueTypeFontProgram font,
        PdfTextShapingOptions? descriptor, out bool covered) {
        covered = false;
        if (!descriptor.HasValue || !PdfTextDiagnostics.TryAnalyzeCompleteManagedRun(text, font,
                descriptor.Value, string.Empty, null, out List<PdfTextEncodingDiagnostic> diagnostics,
                generatedLayout: true)) return false;
        covered = diagnostics.Count == 0;
        return true;
    }

    private static bool TryGetSelectedManagedCoverage(string text, PdfOpenTypeCffFontProgram font,
        PdfTextShapingOptions? descriptor, out bool covered) {
        covered = false;
        if (!descriptor.HasValue || !PdfTextDiagnostics.TryAnalyzeCompleteManagedRun(text, font,
                descriptor.Value, string.Empty, null, out List<PdfTextEncodingDiagnostic> diagnostics,
                generatedLayout: true)) return false;
        covered = diagnostics.Count == 0;
        return true;
    }

}
