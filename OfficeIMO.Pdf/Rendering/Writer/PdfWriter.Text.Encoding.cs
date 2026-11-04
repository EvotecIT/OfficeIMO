using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string EscapeText(string s) => PdfSyntaxEscaper.EscapeLiteralContent(s);

    private static string EncodeWinAnsiHex(string s) {
        var bytes = PdfWinAnsiEncoding.Encode(s);
        const string digits = "0123456789ABCDEF";
        var hex = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++) {
            hex[2 * i] = digits[bytes[i] >> 4];
            hex[2 * i + 1] = digits[bytes[i] & 0xF];
        }
        return new string(hex);
    }

    private static PdfTextShowCommand EncodeTextShowCommand(string text, PdfStandardFont font, PdfCreateOptions? options,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto, double fontMetricScale = 1D, double layoutHorizontalScaling = 100D,
        bool preservesPaintedGlyphs = false, bool requireGlyphClusters = false) {
        string sourceText = text;
        text = WithoutOptionalBreaks(text);
        options?.BeginTextShapingAttempt();
        featureSettings = options?.ResolveTextFeatureSettings(featureSettings) ?? featureSettings;
        PdfTextEncodingDiagnostic? diagnostic = RestoreOptionalBreakDiagnosticIndex(sourceText, GetFirstTextEncodingDiagnostic(text, font, options, featureSettings, textDirection, preservesPaintedGlyphs));
        if (diagnostic != null) {
            throw CreateTextEncodingException(diagnostic, nameof(text));
        }

        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport && !preservesPaintedGlyphs
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            if (options.HasDiagnosticsReport) {
                options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram,
                    shapingOptions: PdfTextDiagnostics.ForEncodingDiagnostics(options, fontProgram.FontName,
                        featureSettings: featureSettings, direction: textDirection, preservesPaintedGlyphs: preservesPaintedGlyphs)));
            }
            PdfTextShapingOptions renderOptions = PdfTextShapingOptions.ForRendering(
                fontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.EffectiveTextShapingLanguage,
                featureSettings,
                textDirection, preservesPaintedGlyphs);
            if (!requireGlyphClusters && !options.UseFontUnitTextLayout && !fontProgram.HasTracking && renderOptions.ShapingProvider == null && renderOptions.FeatureSettings.IsDefault && renderOptions.Direction == OfficeTextDirection.Auto && renderOptions.ShapingMode != PdfTextShapingMode.OpenTypeLigatures && !OfficeUnicodeBidi.ContainsRightToLeft(text)) {
                // The external shaper will not engage, and scalar shaping never positions glyphs, so emit
                // the hex show-string directly without materializing a per-run PdfGlyphRun.
                string glyphHex = PdfUnicodeScalarTextShaper.EncodeGlyphHex(text, fontProgram, renderOptions, out string? actualText, out int advanceWidth1000);
                options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
                return new PdfTextShowCommand(glyphHex, null, actualText, advanceWidth1000: advanceWidth1000, glyphCount: glyphHex.Length / 4);
            }

            PdfGlyphRun glyphRun = fontProgram.ShapeText(text, renderOptions);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
            PdfTextShowCommand command = fontProgram.ToTextShowCommand(text, glyphRun, fontMetricScale);
            return options.UseFontUnitTextLayout ? command.WithFontUnitLayoutAdvance(
                fontProgram.MeasureFontUnitAdvanceWidth1000(glyphRun, layoutHorizontalScaling)) : command;
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport && !preservesPaintedGlyphs
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, cffFontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            if (options.HasDiagnosticsReport) {
                options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cffFontProgram,
                    shapingOptions: PdfTextDiagnostics.ForEncodingDiagnostics(options, cffFontProgram.FontName,
                        featureSettings: featureSettings, direction: textDirection, preservesPaintedGlyphs: preservesPaintedGlyphs)));
            }
            PdfGlyphRun glyphRun = cffFontProgram.ShapeText(text, PdfTextShapingOptions.ForRendering(
                cffFontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.EffectiveTextShapingLanguage,
                featureSettings,
                textDirection, preservesPaintedGlyphs));
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, cffFontProgram.FontName, isOpenTypeCff: true);
            PdfTextShowCommand command = cffFontProgram.ToTextShowCommand(glyphRun);
            return options.UseFontUnitTextLayout ? command.WithFontUnitLayoutAdvance(
                cffFontProgram.MeasureFontUnitAdvanceWidth1000(glyphRun, layoutHorizontalScaling)) : command;
        }

        if (options?.HasDiagnosticsReport == true) {
            options.AddTextShapingDiagnostics(PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text), text, deferProviderCoverable: false);
            // Only a diagnostics report keeps these; without one the scan and its list were discarded.
            options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeWinAnsiText(text));
        }

        int kerningTotal = 0;
        int[]? kerning = options != null && IsKerningEnabled(featureSettings)
            ? GetStandardFontKerning(text, font, out kerningTotal) : null;
        return new PdfTextShowCommand(EncodeWinAnsiHex(text), advanceWidth1000: EstimateSimpleTextWidth(text, font, 1000) + (kerning == null ? 0 : kerningTotal),
            wordSpaceCount: text.Count(character => character == ' '), simpleFontKerning: kerning);
    }

    // Kerning is a default font feature; only an explicit kern=0 disables it.
    private static bool IsKerningEnabled(OfficeTextFeatureSettings? featureSettings) =>
        featureSettings == null || !featureSettings.TryGetValue("kern", out int kerning) || kerning != 0;

    // Standard-font AFM pair kerning, aligned to the single-byte WinAnsi codes of the text.
    private static int[]? GetStandardFontKerning(string text, PdfStandardFont font, out int total) {
        total = 0;
        int[]? kerning = null;
        for (int index = 0; index + 1 < text.Length; index++) {
            int value = PdfStandardFontWidths.GetKerning1000(font, text[index], text[index + 1]);
            if (value == 0) continue;
            if (kerning == null) kerning = new int[text.Length];
            kerning[index] = value;
            total += value;
        }
        return kerning;
    }

    private static double GetStandardFontKerningWidth(string? text, PdfStandardFont font, double fontSize) {
        if (string.IsNullOrEmpty(text)) return 0D;
        GetStandardFontKerning(text!, font, out int total);
        return total * fontSize / 1000D;
    }

    private static PdfTextShowCommand EncodeActualTextAnchor(PdfStandardFont font, PdfCreateOptions options, int count = 1) {
        // The invisible carrier owns the caller's semantic bounds, not shaped paint.
        PdfTextShowCommand command = EncodeTextShowCommand(new string(' ', count), font, options, preservesPaintedGlyphs: true);
        return command.ActualText == null ? command
            : new PdfTextShowCommand(command.GlyphHex, command.PositionedGlyphs, tracking: command.Tracking,
                advanceWidth1000: command.AdvanceWidth1000, wordSpaceCount: command.WordSpaceCount,
                unitsPerEm: command.UnitsPerEm, trackingBoundaries: command.TrackingBoundaries,
                negativeTracking: command.NegativeTracking, fontMetricScale: command.FontMetricScale, glyphCount: command.GlyphCount);
    }

    private static PdfTextShowCommand EncodeTextShowCommand(
        string text,
        PdfStandardFont fallbackFont,
        PdfNamedFontFace? namedFont,
        PdfCreateOptions? options,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto, double fontMetricScale = 1D, double layoutHorizontalScaling = 100D,
        bool preservesPaintedGlyphs = false, bool requireGlyphClusters = false) {
        string sourceText = text;
        text = WithoutOptionalBreaks(text);
        options?.BeginTextShapingAttempt();
        featureSettings = options?.ResolveTextFeatureSettings(featureSettings) ?? featureSettings;
        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = RestoreOptionalBreakDiagnosticIndices(sourceText, PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram,
                    shapingOptions: PdfTextDiagnostics.ForEncodingDiagnostics(options, fontProgram.FontName,
                        featureSettings: featureSettings, direction: textDirection, preservesPaintedGlyphs: preservesPaintedGlyphs)));
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport && !preservesPaintedGlyphs
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            PdfTextShapingOptions renderOptions = PdfTextShapingOptions.ForRendering(
                fontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.EffectiveTextShapingLanguage,
                featureSettings,
                textDirection, preservesPaintedGlyphs);
            if (!requireGlyphClusters && !options.UseFontUnitTextLayout && !fontProgram.HasTracking && renderOptions.ShapingProvider == null && renderOptions.FeatureSettings.IsDefault && renderOptions.Direction == OfficeTextDirection.Auto && renderOptions.ShapingMode != PdfTextShapingMode.OpenTypeLigatures && !OfficeUnicodeBidi.ContainsRightToLeft(text)) {
                // The external shaper will not engage, and scalar shaping never positions glyphs, so emit
                // the hex show-string directly without materializing a per-run PdfGlyphRun.
                string glyphHex = PdfUnicodeScalarTextShaper.EncodeGlyphHex(text, fontProgram, renderOptions, out string? actualText, out int advanceWidth1000);
                options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
                return new PdfTextShowCommand(glyphHex, null, actualText, advanceWidth1000: advanceWidth1000, glyphCount: glyphHex.Length / 4);
            }

            PdfGlyphRun glyphRun = fontProgram.ShapeText(text, renderOptions);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
            PdfTextShowCommand command = fontProgram.ToTextShowCommand(text, glyphRun, fontMetricScale);
            return options.UseFontUnitTextLayout ? command.WithFontUnitLayoutAdvance(
                fontProgram.MeasureFontUnitAdvanceWidth1000(glyphRun, layoutHorizontalScaling)) : command;
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = RestoreOptionalBreakDiagnosticIndices(sourceText, PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cffFontProgram,
                    shapingOptions: PdfTextDiagnostics.ForEncodingDiagnostics(options, cffFontProgram.FontName,
                        featureSettings: featureSettings, direction: textDirection, preservesPaintedGlyphs: preservesPaintedGlyphs)));
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport && !preservesPaintedGlyphs
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, cffFontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            PdfGlyphRun glyphRun = cffFontProgram.ShapeText(text, PdfTextShapingOptions.ForRendering(
                cffFontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.EffectiveTextShapingLanguage,
                featureSettings,
                textDirection, preservesPaintedGlyphs));
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, cffFontProgram.FontName, isOpenTypeCff: true);
            PdfTextShowCommand command = cffFontProgram.ToTextShowCommand(glyphRun);
            return options.UseFontUnitTextLayout ? command.WithFontUnitLayoutAdvance(
                cffFontProgram.MeasureFontUnitAdvanceWidth1000(glyphRun, layoutHorizontalScaling)) : command;
        }

        return EncodeTextShowCommand(sourceText, fallbackFont, options, featureSettings, textDirection, fontMetricScale, layoutHorizontalScaling,
            preservesPaintedGlyphs, requireGlyphClusters);
    }

    internal static IReadOnlyList<PdfTextEncodingDiagnostic> RestoreOptionalBreakDiagnosticIndices(string source, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) =>
        diagnostics.Count == 0 || source.IndexOf('\u200B') < 0 ? diagnostics
            : diagnostics.Select(diagnostic => RestoreOptionalBreakDiagnosticIndex(source, diagnostic)!).ToArray();

    // Glyph indices use normalized text; diagnostics describe the caller's original UTF-16 text.
    private static PdfTextEncodingDiagnostic? RestoreOptionalBreakDiagnosticIndex(string source, PdfTextEncodingDiagnostic? diagnostic) {
        if (diagnostic == null || source.IndexOf('\u200B') < 0) return diagnostic;
        int normalizedIndex = 0;
        int sourceIndex = 0;
        while (sourceIndex < source.Length) {
            if (source[sourceIndex] != '\u200B') {
                if (normalizedIndex == diagnostic.Index) break;
                normalizedIndex++;
            }
            sourceIndex++;
        }
        return diagnostic.WithIndex(sourceIndex);
    }

    private static PdfTextEncodingDiagnostic? GetFirstTextEncodingDiagnostic(string text, PdfStandardFont font, PdfCreateOptions? options, OfficeTextFeatureSettings? featureSettings, OfficeTextDirection direction, bool preservesPaintedGlyphs) {
        System.Collections.Generic.IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = options == null
            ? PdfTextDiagnostics.AnalyzeWinAnsiText(text, "generated text")
            : PdfTextDiagnostics.AnalyzeGeneratedEncoding(text, options, font, "generated text", featureSettings, direction, preservesPaintedGlyphs);

        return diagnostics.Count == 0 ? null : diagnostics[0];
    }

    private static ArgumentException CreateTextEncodingException(PdfTextEncodingDiagnostic diagnostic, string paramName) {
        var exception = new ArgumentException(diagnostic.Message, paramName);
        exception.Data["code"] = diagnostic.Code;
        exception.Data["source"] = diagnostic.Source;
        exception.Data["index"] = diagnostic.Index;
        exception.Data["codePoint"] = diagnostic.CodePoint;
        exception.Data["text"] = diagnostic.Text;
        exception.Data["isControlCharacter"] = diagnostic.IsControlCharacter;
        if (!string.IsNullOrWhiteSpace(diagnostic.Location)) {
            exception.Data["location"] = diagnostic.Location;
        }

        if (!string.IsNullOrWhiteSpace(diagnostic.Encoding)) {
            exception.Data["encoding"] = diagnostic.Encoding;
        }

        if (!string.IsNullOrWhiteSpace(diagnostic.Remediation)) {
            exception.Data["remediation"] = diagnostic.Remediation;
        }

        return exception;
    }

    private static int GetScalarUtf16Length(string text, int index) {
        if (index < 0 || index >= text.Length) {
            throw new ArgumentOutOfRangeException(nameof(index), "Text scalar index must be inside the string.");
        }

        return char.IsHighSurrogate(text[index]) &&
            index + 1 < text.Length &&
            char.IsLowSurrogate(text[index + 1])
                ? 2
                : 1;
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
    }
}
