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

    private static PdfTextShowCommand EncodeTextShowCommand(string text, PdfStandardFont font, PdfOptions? options,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) {
        options?.BeginTextShapingAttempt();
        PdfTextEncodingDiagnostic? diagnostic = GetFirstTextEncodingDiagnostic(text, font, options);
        if (diagnostic != null) {
            throw CreateTextEncodingException(diagnostic, nameof(text));
        }

        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            if (options.HasDiagnosticsReport) {
                options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram));
            }
            PdfTextShapingOptions renderOptions = PdfTextShapingOptions.ForRendering(
                fontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.Language,
                featureSettings,
                textDirection);
            if (renderOptions.ShapingProvider == null && renderOptions.FeatureSettings.IsDefault && renderOptions.Direction == OfficeTextDirection.Auto && renderOptions.ShapingMode != PdfTextShapingMode.OpenTypeLigatures) {
                // The external shaper will not engage, and scalar shaping never positions glyphs, so emit
                // the hex show-string directly without materializing a per-run PdfGlyphRun.
                string glyphHex = PdfUnicodeScalarTextShaper.EncodeGlyphHex(text, fontProgram, renderOptions, out string? actualText, out int advanceWidth1000);
                options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
                return new PdfTextShowCommand(glyphHex, null, actualText, advanceWidth1000: advanceWidth1000, glyphCount: glyphHex.Length / 4);
            }

            PdfGlyphRun glyphRun = fontProgram.ShapeText(text, renderOptions);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
            return glyphRun.ToTextShowCommand();
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, cffFontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            if (options.HasDiagnosticsReport) {
                options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cffFontProgram));
            }
            PdfGlyphRun glyphRun = cffFontProgram.ShapeText(text, PdfTextShapingOptions.ForRendering(
                cffFontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.Language,
                featureSettings,
                textDirection));
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, cffFontProgram.FontName, isOpenTypeCff: true);
            return glyphRun.ToTextShowCommand();
        }

        if (options?.HasDiagnosticsReport == true) {
            options.AddTextShapingDiagnostics(PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text), text, deferProviderCoverable: false);
            // Only a diagnostics report keeps these; without one the scan and its list were discarded.
            options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeWinAnsiText(text));
        }

        return new PdfTextShowCommand(EncodeWinAnsiHex(text), advanceWidth1000: EstimateSimpleTextWidth(text, font, 1000),
            wordSpaceCount: text.Count(character => character == ' '));
    }

    private static PdfTextShowCommand EncodeActualTextAnchor(PdfStandardFont font, PdfOptions options, int count = 1) {
        PdfTextShowCommand command = EncodeTextShowCommand(new string(' ', count), font, options);
        return new PdfTextShowCommand(command.GlyphHex, command.PositionedGlyphs,
            advanceWidth1000: command.AdvanceWidth1000, wordSpaceCount: command.WordSpaceCount);
    }

    private static PdfTextShowCommand EncodeTextShowCommand(
        string text,
        PdfStandardFont fallbackFont,
        PdfNamedFontFace? namedFont,
        PdfOptions? options,
        OfficeTextFeatureSettings? featureSettings = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) {
        options?.BeginTextShapingAttempt();
        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram);
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            PdfTextShapingOptions renderOptions = PdfTextShapingOptions.ForRendering(
                fontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.Language,
                featureSettings,
                textDirection);
            if (renderOptions.ShapingProvider == null && renderOptions.FeatureSettings.IsDefault && renderOptions.Direction == OfficeTextDirection.Auto && renderOptions.ShapingMode != PdfTextShapingMode.OpenTypeLigatures) {
                // The external shaper will not engage, and scalar shaping never positions glyphs, so emit
                // the hex show-string directly without materializing a per-run PdfGlyphRun.
                string glyphHex = PdfUnicodeScalarTextShaper.EncodeGlyphHex(text, fontProgram, renderOptions, out string? actualText, out int advanceWidth1000);
                options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
                return new PdfTextShowCommand(glyphHex, null, actualText, advanceWidth1000: advanceWidth1000, glyphCount: glyphHex.Length / 4);
            }

            PdfGlyphRun glyphRun = fontProgram.ShapeText(text, renderOptions);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
            return glyphRun.ToTextShowCommand();
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cffFontProgram);
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, cffFontProgram, featureSettings: featureSettings)
                : Array.Empty<PdfTextShapingDiagnostic>();
            PdfGlyphRun glyphRun = cffFontProgram.ShapeText(text, PdfTextShapingOptions.ForRendering(
                cffFontProgram.FontName,
                options.TextShapingModeSnapshot,
                options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate,
                options.Language,
                featureSettings,
                textDirection));
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, cffFontProgram.FontName, isOpenTypeCff: true);
            return glyphRun.ToTextShowCommand();
        }

        return EncodeTextShowCommand(text, fallbackFont, options, featureSettings, textDirection);
    }

    private static PdfTextEncodingDiagnostic? GetFirstTextEncodingDiagnostic(string text, PdfStandardFont font, PdfOptions? options) {
        System.Collections.Generic.IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = options == null
            ? PdfTextDiagnostics.AnalyzeWinAnsiText(text, "generated text")
            : PdfTextDiagnostics.AnalyzeGeneratedText(text, options, font, "generated text");

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
