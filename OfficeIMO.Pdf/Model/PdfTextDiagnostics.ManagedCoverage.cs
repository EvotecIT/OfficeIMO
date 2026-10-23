using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextDiagnostics {
    /// <summary>Reuses the current writer descriptor without recording subset usage or provider coverage.</summary>
    internal static PdfTextShapingOptions ForEncodingDiagnostics(PdfOptions options, string fontName,
        string source = "", OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection direction = OfficeTextDirection.Auto, bool preservesPaintedGlyphs = false) =>
        new(recordGlyphUsage: false, throwOnMissingGlyph: false, skipLayoutControls: true,
            reportControlCharacters: true, source: source, fontName: fontName,
            shapingMode: options.TextShapingModeSnapshot, shapingProvider: options.TextShapingProviderSnapshot,
            language: options.EffectiveTextShapingLanguage,
            featureSettings: options.ResolveTextFeatureSettings(featureSettings), direction: direction,
            preservesPaintedGlyphs: preservesPaintedGlyphs);

    internal static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeGeneratedEncoding(string text, PdfOptions options,
        PdfStandardFont font, string source, OfficeTextFeatureSettings? featureSettings,
        OfficeTextDirection direction, bool preservesPaintedGlyphs) {
        if (!PdfExternalTextShaper.RequiresCompleteDefaultLatinShaping(text))
            return AnalyzeGeneratedText(text, options, font, source);
        // Complete normalization and retained source controls need the effective
        // writer descriptor. Reuse the existing generated-run font resolver.
        PdfTextRun run = new PdfTextRun(text, font: font)
            .WithFeatureSettings(options.ResolveTextFeatureSettings(featureSettings))
            .WithTextDirection(direction).WithPaintedGlyphPreservation(preservesPaintedGlyphs);
        return AnalyzeGeneratedTextRuns(new[] { run }, options, font, source);
    }

    internal static bool TryAnalyzeCompleteManagedRun(string text, PdfTrueTypeFontProgram font,
        PdfTextShapingOptions options, string location, int? runIndex, out List<PdfTextEncodingDiagnostic> diagnostics,
        bool generatedLayout = false) {
        diagnostics = null!;
        if (generatedLayout && HasManagedLayoutSeparator(text)
            && TryAnalyzeManagedLayoutSegments(text, options,
                segment => AnalyzeEmbeddedFontText(segment, font, options.Source, location, runIndex,
                    options.ShapingMode, options), out diagnostics)) return true;
        if (!CanAnalyzeCompleteManagedRun(text, options)
            || !PdfExternalTextShaper.TryShapeText(text, font, options, out PdfGlyphRun run)
            || run.SourceShapingResult == null && run.Diagnostics.Count == 0) return false;
        // A declined complete run does not cover ordinary missing glyphs. Reuse
        // the scalar diagnostic owner without features, providers or subset usage.
        List<PdfTextEncodingDiagnostic>? raw = run.SourceShapingResult == null
            ? AnalyzeEmbeddedFontTextCore(text, font, options.Source, options.FontName,
                PdfTextShapingMode.UnicodeScalar)
            : null;
        diagnostics = CollectManagedCoverageDiagnostics(run, raw, location, runIndex);
        return true;
    }

    internal static bool TryAnalyzeCompleteManagedRun(string text, PdfOpenTypeCffFontProgram font,
        PdfTextShapingOptions options, string location, int? runIndex, out List<PdfTextEncodingDiagnostic> diagnostics,
        bool generatedLayout = false) {
        diagnostics = null!;
        if (generatedLayout && HasManagedLayoutSeparator(text)
            && TryAnalyzeManagedLayoutSegments(text, options,
                segment => AnalyzeEmbeddedFontText(segment, font, options.Source, location, runIndex,
                    options.ShapingMode, options), out diagnostics)) return true;
        if (!CanAnalyzeCompleteManagedRun(text, options)
            || !PdfExternalTextShaper.TryShapeText(text, font, options, out PdfGlyphRun run)
            || run.SourceShapingResult == null && run.Diagnostics.Count == 0) return false;
        // A declined complete run does not cover ordinary missing glyphs. Reuse
        // the scalar diagnostic owner without features, providers or subset usage.
        List<PdfTextEncodingDiagnostic>? raw = run.SourceShapingResult == null
            ? AnalyzeEmbeddedFontTextCore(text, font, options.Source, options.FontName,
                PdfTextShapingMode.UnicodeScalar)
            : null;
        diagnostics = CollectManagedCoverageDiagnostics(run, raw, location, runIndex);
        return true;
    }

    // These layout separators are consumed before the writer paints a run.
    // Keep their original UTF-16 offsets while reusing the generated diagnostic
    // owner for each painted segment. Raw/public scalar analysis stays separate.
    private static bool TryAnalyzeManagedLayoutSegments(string text, PdfTextShapingOptions options,
        Func<string, List<PdfTextEncodingDiagnostic>> analyzeSegment,
        out List<PdfTextEncodingDiagnostic> diagnostics) {
        diagnostics = null!;
        if (!CanUseManagedCoverageDescriptor(options)
            || !PdfExternalTextShaper.RequiresCompleteDefaultLatinShaping(text)) return false;
        var result = new List<PdfTextEncodingDiagnostic>();
        int start = 0, separators = 0;
        for (int index = 0; index <= text.Length; index++) {
            if (index < text.Length && !PdfWriter.IsFallbackLayoutSeparator(text[index])) continue;
            if (index > start) {
                string segment = text.Substring(start, index - start);
                foreach (PdfTextEncodingDiagnostic diagnostic in analyzeSegment(segment))
                    result.Add(diagnostic.WithIndex(checked(start + diagnostic.Index)));
            }
            if (index == text.Length) break;
            if (++separators > PdfWriter.MaximumTextLayoutLines)
                throw new System.IO.InvalidDataException("PDF text exceeds the 100,000-layout-control limit.");
            if (text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            start = index + 1;
        }
        diagnostics = result;
        return true;
    }

    private static bool HasManagedLayoutSeparator(string text) =>
        text.Contains('\n') || text.Contains('\r') || text.Contains('\t');

    private static bool CanUseManagedCoverageDescriptor(PdfTextShapingOptions options) =>
        !options.RecordGlyphUsage && !options.ThrowOnMissingGlyph && !options.PreservesPaintedGlyphs
        && options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures
        && (options.ShapingProvider == null
            || ReferenceEquals(options.ShapingProvider, OfficeManagedTextShapingProvider.Instance));

    private static bool CanAnalyzeCompleteManagedRun(string text, PdfTextShapingOptions options) {
        if (!CanUseManagedCoverageDescriptor(options)
            || !PdfExternalTextShaper.RequiresCompleteDefaultLatinShaping(text)) return false;
        // C0/DEL and layout controls retain the existing scalar diagnostic owner.
        // This proof admits one complete managed font run; it never erases those warnings.
        foreach (char character in text) if (character < ' ' || character == '\u007F') return false;
        return true;
    }

    private static List<PdfTextEncodingDiagnostic> CollectManagedCoverageDiagnostics(PdfGlyphRun run,
        List<PdfTextEncodingDiagnostic>? raw, string location, int? runIndex) {
        // Successful complete shaping alone replaces raw cmap coverage. Preserve
        // the original diagnostic identity while applying document annotations.
        if (raw == null)
            return run.Diagnostics.Select(diagnostic => diagnostic.WithRunLocation(location, runIndex)).ToList();
        var byIndex = new SortedDictionary<int, PdfTextEncodingDiagnostic>();
        foreach (PdfTextEncodingDiagnostic diagnostic in raw) byIndex[diagnostic.Index] = diagnostic;
        // Both owners can reject the same source character. Keep an explicit
        // blank/source failure and retain every other missing-glyph diagnostic.
        foreach (PdfTextEncodingDiagnostic diagnostic in run.Diagnostics) byIndex[diagnostic.Index] = diagnostic;
        return byIndex.Values.Select(diagnostic => diagnostic.WithRunLocation(location, runIndex)).ToList();
    }
}
