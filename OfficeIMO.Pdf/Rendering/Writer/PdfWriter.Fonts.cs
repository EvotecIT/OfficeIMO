using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly string[] LayoutLineSeparators = { "\r\n", "\r", "\n" };

    private static PdfStandardFont ChooseNormal(PdfStandardFont requested) => requested switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaBoldOblique => PdfStandardFont.Helvetica,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold or PdfStandardFont.TimesBoldItalic => PdfStandardFont.TimesRoman,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold or PdfStandardFont.CourierBoldOblique => PdfStandardFont.Courier,
        _ => ThrowUnsupportedStandardFont(requested)
    };

    private static PdfStandardFont ChooseBold(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaBoldOblique => PdfStandardFont.HelveticaBold,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold or PdfStandardFont.TimesBoldItalic => PdfStandardFont.TimesBold,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold or PdfStandardFont.CourierBoldOblique => PdfStandardFont.CourierBold,
        _ => ThrowUnsupportedStandardFont(normal)
    };

    private static PdfStandardFont ChooseItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaBoldOblique => PdfStandardFont.HelveticaOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold or PdfStandardFont.TimesBoldItalic => PdfStandardFont.TimesItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold or PdfStandardFont.CourierBoldOblique => PdfStandardFont.CourierOblique,
        _ => ThrowUnsupportedStandardFont(normal)
    };

    private static PdfStandardFont ChooseBoldItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaBoldOblique => PdfStandardFont.HelveticaBoldOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold or PdfStandardFont.TimesBoldItalic => PdfStandardFont.TimesBoldItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold or PdfStandardFont.CourierBoldOblique => PdfStandardFont.CourierBoldOblique,
        _ => ThrowUnsupportedStandardFont(normal)
    };

    private static string GetStandardFontResourceName(PdfStandardFont font, PdfStandardFont defaultNormalFont) {
        if (font == defaultNormalFont) return "F1";
        if (font == ChooseBold(defaultNormalFont)) return "F2";
        if (font == ChooseItalic(defaultNormalFont)) return "F3";
        if (font == ChooseBoldItalic(defaultNormalFont)) return "F4";

        return GetIndependentStandardFontResourceName(font);
    }

    private static string GetFontResourceName(PdfStandardFont fallbackFont, PdfNamedFontFace? namedFont, PdfStandardFont defaultNormalFont) =>
        namedFont?.ResourceName ?? GetStandardFontResourceName(fallbackFont, defaultNormalFont);

    private static string GetIndependentStandardFontResourceName(PdfStandardFont font) => font switch {
        PdfStandardFont.Helvetica => "F11",
        PdfStandardFont.HelveticaBold => "F12",
        PdfStandardFont.HelveticaOblique => "F13",
        PdfStandardFont.HelveticaBoldOblique => "F14",
        PdfStandardFont.TimesRoman => "F15",
        PdfStandardFont.TimesBold => "F16",
        PdfStandardFont.TimesItalic => "F17",
        PdfStandardFont.TimesBoldItalic => "F18",
        PdfStandardFont.Courier => "F19",
        PdfStandardFont.CourierBold => "F20",
        PdfStandardFont.CourierOblique => "F21",
        PdfStandardFont.CourierBoldOblique => "F22",
        _ => ThrowUnsupportedStandardFontResource(font)
    };

    private static PdfStandardFont ResolveFontFromResourceName(string resourceName, PdfStandardFont defaultNormalFont) {
        string name = resourceName != null && resourceName.Length > 0 && resourceName[0] == '/'
            ? resourceName.Substring(1)
            : resourceName ?? string.Empty;

        switch (name) {
            case "F1":
                return defaultNormalFont;
            case "F2":
                return ChooseBold(defaultNormalFont);
            case "F3":
                return ChooseItalic(defaultNormalFont);
            case "F4":
                return ChooseBoldItalic(defaultNormalFont);
            case "F11":
                return PdfStandardFont.Helvetica;
            case "F12":
                return PdfStandardFont.HelveticaBold;
            case "F13":
                return PdfStandardFont.HelveticaOblique;
            case "F14":
                return PdfStandardFont.HelveticaBoldOblique;
            case "F15":
                return PdfStandardFont.TimesRoman;
            case "F16":
                return PdfStandardFont.TimesBold;
            case "F17":
                return PdfStandardFont.TimesItalic;
            case "F18":
                return PdfStandardFont.TimesBoldItalic;
            case "F19":
                return PdfStandardFont.Courier;
            case "F20":
                return PdfStandardFont.CourierBold;
            case "F21":
                return PdfStandardFont.CourierOblique;
            case "F22":
                return PdfStandardFont.CourierBoldOblique;
            default:
                return defaultNormalFont;
        }
    }

    private static double GlyphWidthEmFor(PdfStandardFont font) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => 0.6,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => 0.55,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => 0.5,
        _ => ThrowUnsupportedStandardFontWidth(font)
    };

    private static double SpaceWidthEmFor(PdfStandardFont font) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => 0.6,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => 0.278,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => 0.25,
        _ => ThrowUnsupportedStandardFontWidth(font)
    };

    // A non-allocating '\r'/'\n' scan. `text.Any(c => ...)` allocates a CharEnumerator per call, and this
    // runs once per token/cell/paragraph during layout. string.Contains(char) is unavailable on
    // netstandard2.0/net472 and string.IndexOf(char) trips CA2249-as-error, so use an index loop.
    private static bool ContainsLineBreak(string text) {
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\r' || text[i] == '\n') {
                return true;
            }
        }

        return false;
    }

    internal static double EstimateSimpleTextWidth(string? text, PdfStandardFont font, double fontSize) {
        if (string.IsNullOrEmpty(text)) {
            return 0;
        }

        double width = 0;
        for (int i = 0; i < text!.Length; i++) {
            width += StandardGlyphWidthEmFor(font, text[i]) * fontSize;
        }

        return width;
    }

    // The multi-line branch is factored out so its lambda-free loop keeps the closure that would capture
    // these parameters off the single-line hot path — this method is called per token during wrapping,
    // and a captured lambda forces a display-class allocation on every call even when unused.
    private static double MaxLineWidthForOptions(string text, PdfStandardFont font, double fontSize, PdfOptions? options, OfficeTextFeatureSettings? featureSettings) {
        // Seeded below the possible range (not 0) so the result equals the original LINQ .Max() exactly;
        // Split with a line break present always yields at least one element, so max is always assigned.
        double max = double.NegativeInfinity;
        foreach (string line in text.Split(LayoutLineSeparators, StringSplitOptions.None)) {
            double width = EstimateSimpleTextWidthForOptions(line, font, fontSize, options, featureSettings);
            if (width > max) max = width;
        }

        return max;
    }

    private static double EstimateSimpleTextWidthForOptions(string? text, PdfStandardFont font, double fontSize, PdfOptions? options, OfficeTextFeatureSettings? featureSettings = null) {
        if (!string.IsNullOrEmpty(text) && ContainsLineBreak(text!)) {
            return MaxLineWidthForOptions(text!, font, fontSize, options, featureSettings);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            string value = text ?? string.Empty;
            if (options.HasDiagnosticsReport) {
                options.AddTextShapingDiagnostics(
                    PdfTextDiagnostics.AnalyzeAdvancedTextLayout(value, fontProgram),
                    value,
                    deferProviderCoverable: options.TextShapingProviderSnapshot != null);
            }

            // Only run the allocating embedded-font diagnostic scan when a diagnostics report needs it;
            // MeasureTextWidth already throws on a missing glyph, so validation is preserved regardless.
            if (options.HasDiagnosticsReport) {
                IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, fontProgram);
                options.AddTextDiagnostics(diagnostics);
                if (diagnostics.Count > 0) {
                    throw CreateTextEncodingException(diagnostics[0], nameof(text));
                }
            }

            return fontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language, featureSettings);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            string value = text ?? string.Empty;
            if (options.HasDiagnosticsReport) {
                options.AddTextShapingDiagnostics(
                    PdfTextDiagnostics.AnalyzeAdvancedTextLayout(value, cffFontProgram),
                    value,
                    deferProviderCoverable: options.TextShapingProviderSnapshot != null);
            }

            // Only run the allocating embedded-font diagnostic scan when a diagnostics report needs it;
            // MeasureTextWidth already throws on a missing glyph, so validation is preserved regardless.
            if (options.HasDiagnosticsReport) {
                IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, cffFontProgram);
                options.AddTextDiagnostics(diagnostics);
                if (diagnostics.Count > 0) {
                    throw CreateTextEncodingException(diagnostics[0], nameof(text));
                }
            }

            return cffFontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language, featureSettings);
        }

        return EstimateSimpleTextWidth(text, font, fontSize);
    }

    private static double MaxLineWidthForOptions(string text, PdfStandardFont fallbackFont, PdfNamedFontFace? namedFont, double fontSize, PdfOptions? options, OfficeTextFeatureSettings? featureSettings) {
        double max = double.NegativeInfinity;
        foreach (string line in text.Split(LayoutLineSeparators, StringSplitOptions.None)) {
            double width = EstimateSimpleTextWidthForOptions(line, fallbackFont, namedFont, fontSize, options, featureSettings);
            if (width > max) max = width;
        }

        return max;
    }

    private static double EstimateSimpleTextWidthForOptions(
        string? text,
        PdfStandardFont fallbackFont,
        PdfNamedFontFace? namedFont,
        double fontSize,
        PdfOptions? options,
        OfficeTextFeatureSettings? featureSettings = null) {
        if (!string.IsNullOrEmpty(text) && ContainsLineBreak(text!)) {
            return MaxLineWidthForOptions(text!, fallbackFont, namedFont, fontSize, options, featureSettings);
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            string value = text ?? string.Empty;
            if (options.HasDiagnosticsReport) {
                options.AddTextShapingDiagnostics(
                    PdfTextDiagnostics.AnalyzeAdvancedTextLayout(value, fontProgram),
                    value,
                    deferProviderCoverable: options.TextShapingProviderSnapshot != null);
            }

            // Only run the allocating embedded-font diagnostic scan when a diagnostics report needs it;
            // MeasureTextWidth already throws on a missing glyph, so validation is preserved regardless.
            if (options.HasDiagnosticsReport) {
                IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, fontProgram);
                options.AddTextDiagnostics(diagnostics);
                if (diagnostics.Count > 0) {
                    throw CreateTextEncodingException(diagnostics[0], nameof(text));
                }
            }

            return fontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language, featureSettings);
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            string value = text ?? string.Empty;
            if (options.HasDiagnosticsReport) {
                options.AddTextShapingDiagnostics(
                    PdfTextDiagnostics.AnalyzeAdvancedTextLayout(value, cffFontProgram),
                    value,
                    deferProviderCoverable: options.TextShapingProviderSnapshot != null);
            }

            // Only run the allocating embedded-font diagnostic scan when a diagnostics report needs it;
            // MeasureTextWidth already throws on a missing glyph, so validation is preserved regardless.
            if (options.HasDiagnosticsReport) {
                IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, cffFontProgram);
                options.AddTextDiagnostics(diagnostics);
                if (diagnostics.Count > 0) {
                    throw CreateTextEncodingException(diagnostics[0], nameof(text));
                }
            }

            return cffFontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language, featureSettings);
        }

        return EstimateSimpleTextWidthForOptions(text, fallbackFont, fontSize, options, featureSettings);
    }

    internal static double EstimateSimpleTextWidth1000(string? text, PdfStandardFont font) =>
        EstimateSimpleTextWidth(text, font, 1000D);

    internal static bool TryGetStandardFontByBaseFontName(string? baseFontName, out PdfStandardFont font) {
        font = PdfStandardFont.Helvetica;
        if (string.IsNullOrWhiteSpace(baseFontName)) {
            return false;
        }

        string normalized = StripSubsetPrefix(baseFontName!);
        if (EqualsIgnoreCase(normalized, "Helvetica-BoldOblique")) {
            font = PdfStandardFont.HelveticaBoldOblique;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Helvetica-Bold")) {
            font = PdfStandardFont.HelveticaBold;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Helvetica-Oblique")) {
            font = PdfStandardFont.HelveticaOblique;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Helvetica")) {
            font = PdfStandardFont.Helvetica;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Times-BoldItalic")) {
            font = PdfStandardFont.TimesBoldItalic;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Times-Bold")) {
            font = PdfStandardFont.TimesBold;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Times-Italic")) {
            font = PdfStandardFont.TimesItalic;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Times-Roman")) {
            font = PdfStandardFont.TimesRoman;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Courier-BoldOblique")) {
            font = PdfStandardFont.CourierBoldOblique;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Courier-Bold")) {
            font = PdfStandardFont.CourierBold;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Courier-Oblique")) {
            font = PdfStandardFont.CourierOblique;
            return true;
        }

        if (EqualsIgnoreCase(normalized, "Courier")) {
            font = PdfStandardFont.Courier;
            return true;
        }

        return false;
    }

    private static string StripSubsetPrefix(string baseFontName) {
        int plusIndex = baseFontName.IndexOf('+');
        if (plusIndex > 0 && plusIndex < baseFontName.Length - 1) {
            return baseFontName.Substring(plusIndex + 1);
        }

        return baseFontName;
    }

    private static bool EqualsIgnoreCase(string left, string right) =>
        string.Equals(left, right, System.StringComparison.OrdinalIgnoreCase);

    private static double StandardGlyphWidthEmFor(PdfStandardFont font, char value) {
        if (font == PdfStandardFont.Courier ||
            font == PdfStandardFont.CourierBold ||
            font == PdfStandardFont.CourierOblique ||
            font == PdfStandardFont.CourierBoldOblique) {
            return 0.6;
        }

        if (PdfStandardFontWidths.TryGetWidth1000(font, value, out int width)) {
            return width / 1000D;
        }

        // An accented letter outside the core glyph set is measured by its base letter.
        if (value > '~' && char.IsLetter(value)) {
            string decomposed = value.ToString().Normalize(System.Text.NormalizationForm.FormD);
            if (decomposed.Length > 1 && decomposed[0] != value &&
                PdfStandardFontWidths.TryGetWidth1000(font, decomposed[0], out width)) {
                return width / 1000D;
            }
        }

        return GlyphWidthEmFor(font);
    }

    private static double GetDescender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => fontSize * 0.23,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => fontSize * 0.22,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => fontSize * 0.26,
        _ => ThrowUnsupportedStandardFontWidth(font)
    };

    private static double GetDescenderForOptions(PdfStandardFont font, double fontSize, PdfOptions? options) {
        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return fontProgram.GetDescender(fontSize);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return cffFontProgram.GetDescender(fontSize);
        }

        return GetDescender(font, fontSize);
    }

    private static double GetDescenderForOptions(PdfStandardFont font, PdfNamedFontFace? namedFont, double fontSize, PdfOptions? options) {
        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return fontProgram.GetDescender(fontSize);
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return cffFontProgram.GetDescender(fontSize);
        }

        return GetDescenderForOptions(font, fontSize, options);
    }

    private static double GetAscender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => fontSize * 0.72,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => fontSize * 0.74,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => fontSize * 0.72,
        _ => ThrowUnsupportedStandardFontWidth(font)
    };

    private static double GetAscenderForOptions(PdfStandardFont font, double fontSize, PdfOptions? options) {
        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return fontProgram.GetAscender(fontSize);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return cffFontProgram.GetAscender(fontSize);
        }

        return GetAscender(font, fontSize);
    }

    private static double GetAscenderForOptions(PdfStandardFont font, PdfNamedFontFace? namedFont, double fontSize, PdfOptions? options) {
        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return fontProgram.GetAscender(fontSize);
        }

        if (namedFont.HasValue &&
            options != null &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return cffFontProgram.GetAscender(fontSize);
        }

        return GetAscenderForOptions(font, fontSize, options);
    }

    private static PdfStandardFont ThrowUnsupportedStandardFont(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        throw new System.ArgumentOutOfRangeException(nameof(font), "PDF font must be one of the supported standard PDF fonts.");
    }

    private static double ThrowUnsupportedStandardFontWidth(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        throw new System.ArgumentOutOfRangeException(nameof(font), "PDF font must be one of the supported standard PDF fonts.");
    }

    private static string ThrowUnsupportedStandardFontResource(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        throw new System.ArgumentOutOfRangeException(nameof(font), "PDF font must be one of the supported standard PDF fonts.");
    }
}
