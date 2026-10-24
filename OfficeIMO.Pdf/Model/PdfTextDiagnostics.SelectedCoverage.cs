namespace OfficeIMO.Pdf;

internal static partial class PdfTextDiagnostics {
    // Preserve selected caller fonts before raw fallback planning. Candidate
    // ranges and profile-family priority remain in the existing fallback owner.
    private static List<PdfTextEncodingDiagnostic> AnalyzeGeneratedSelectedFont(string text,
        PdfTrueTypeFontProgram font, string source, string location, int? runIndex,
        PdfTextShapingMode shapingMode, PdfEmbeddedFontFallbackSet? fallbackSet,
        PdfTextShapingOptions descriptor, bool preferSelected) {
        if (fallbackSet == null)
            return AnalyzeEmbeddedFontText(text, font, source, location, runIndex, shapingMode, descriptor);
        PdfTextShapingOptions? coverage = preferSelected ? descriptor : null;
        if (coverage.HasValue && TryAnalyzeCompleteManagedRun(text, font, coverage.Value,
                location, runIndex, out List<PdfTextEncodingDiagnostic> complete, generatedLayout: true)
            && complete.Count == 0) return complete;
        return AnalyzeGeneratedTextWithFallback(text, fallbackSet, source, location, runIndex, shapingMode,
            (string value, int index, out int length) => TryGetCoveredTextElementLength(value, index,
                font, shapingMode, out length, coverage));
    }

    // Preserve selected caller fonts before raw fallback planning. Candidate
    // ranges and profile-family priority remain in the existing fallback owner.
    private static List<PdfTextEncodingDiagnostic> AnalyzeGeneratedSelectedFont(string text,
        PdfOpenTypeCffFontProgram font, string source, string location, int? runIndex,
        PdfTextShapingMode shapingMode, PdfEmbeddedFontFallbackSet? fallbackSet,
        PdfTextShapingOptions descriptor, bool preferSelected) {
        if (fallbackSet == null)
            return AnalyzeEmbeddedFontText(text, font, source, location, runIndex, shapingMode, descriptor);
        PdfTextShapingOptions? coverage = preferSelected ? descriptor : null;
        if (coverage.HasValue && TryAnalyzeCompleteManagedRun(text, font, coverage.Value,
                location, runIndex, out List<PdfTextEncodingDiagnostic> complete, generatedLayout: true)
            && complete.Count == 0) return complete;
        return AnalyzeGeneratedTextWithFallback(text, fallbackSet, source, location, runIndex, shapingMode,
            (string value, int index, out int length) => TryGetCoveredTextElementLength(value, index,
                font, shapingMode, out length, coverage));
    }

}
