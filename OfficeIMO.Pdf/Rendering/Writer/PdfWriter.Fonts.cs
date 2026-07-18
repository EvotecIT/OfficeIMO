namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly string[] LayoutLineSeparators = { "\r\n", "\r", "\n" };

    private static readonly int[] HelveticaBoldAsciiWidths = new[] {
        278, 333, 474, 556, 556, 889, 722, 238, 333, 333,
        389, 584, 278, 333, 278, 278, 556, 556, 556, 556,
        556, 556, 556, 556, 556, 556, 333, 333, 584, 584,
        584, 611, 975, 722, 722, 722, 722, 667, 611, 778,
        722, 278, 556, 722, 611, 833, 722, 778, 667, 778,
        722, 667, 611, 722, 667, 944, 667, 667, 611, 333,
        278, 333, 584, 556, 333, 556, 611, 556, 611, 556,
        333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
        611, 611, 389, 556, 333, 611, 556, 778, 556, 556,
        500, 389, 280, 389, 584
    };

    private static readonly int[] TimesBoldAsciiWidths = new[] {
        250, 333, 555, 500, 500, 1000, 833, 278, 333, 333,
        500, 570, 250, 333, 250, 278, 500, 500, 500, 500,
        500, 500, 500, 500, 500, 500, 333, 333, 570, 570,
        570, 500, 930, 722, 667, 722, 722, 667, 611, 778,
        778, 389, 500, 778, 667, 944, 722, 778, 611, 778,
        722, 556, 667, 722, 722, 1000, 722, 722, 667, 333,
        278, 333, 581, 500, 333, 500, 556, 444, 556, 444,
        333, 500, 556, 278, 333, 556, 278, 833, 556, 500,
        556, 556, 444, 389, 333, 556, 500, 722, 500, 500,
        444, 394, 220, 394, 520
    };

    private static readonly int[] TimesItalicAsciiWidths = new[] {
        250, 333, 420, 500, 500, 833, 778, 214, 333, 333,
        500, 675, 250, 333, 250, 278, 500, 500, 500, 500,
        500, 500, 500, 500, 500, 500, 333, 333, 675, 675,
        675, 500, 920, 611, 611, 667, 722, 611, 611, 722,
        722, 333, 444, 667, 556, 833, 667, 722, 611, 722,
        611, 500, 556, 722, 611, 833, 611, 556, 556, 389,
        278, 389, 422, 500, 333, 500, 500, 444, 500, 444,
        278, 500, 500, 278, 278, 444, 278, 722, 500, 500,
        500, 500, 389, 389, 278, 500, 444, 667, 444, 444,
        389, 400, 275, 400, 541
    };

    private static readonly int[] TimesBoldItalicAsciiWidths = new[] {
        250, 389, 555, 500, 500, 833, 778, 278, 333, 333,
        500, 570, 250, 333, 250, 278, 500, 500, 500, 500,
        500, 500, 500, 500, 500, 500, 333, 333, 570, 570,
        570, 500, 832, 667, 667, 667, 722, 667, 667, 722,
        778, 389, 500, 667, 611, 889, 722, 722, 611, 722,
        667, 556, 611, 722, 667, 889, 667, 611, 611, 333,
        278, 333, 570, 500, 333, 500, 500, 444, 500, 444,
        333, 500, 556, 278, 278, 500, 278, 778, 556, 500,
        500, 500, 389, 389, 278, 556, 444, 667, 500, 444,
        389, 348, 220, 348, 570
    };

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

    private static double EstimateSimpleTextWidthForOptions(string? text, PdfStandardFont font, double fontSize, PdfOptions? options) {
        if (!string.IsNullOrEmpty(text) && text!.Any(character => character == '\r' || character == '\n')) {
            string layoutText = text!;
            return layoutText
                .Split(LayoutLineSeparators, StringSplitOptions.None)
                .Max(line => EstimateSimpleTextWidthForOptions(line, font, fontSize, options));
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

            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, fontProgram);
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            return fontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language);
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

            IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeEmbeddedFontText(value, cffFontProgram);
            options.AddTextDiagnostics(diagnostics);
            if (diagnostics.Count > 0) {
                throw CreateTextEncodingException(diagnostics[0], nameof(text));
            }

            return cffFontProgram.MeasureTextWidth(text, fontSize, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.Language);
        }

        return EstimateSimpleTextWidth(text, font, fontSize);
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

        if (font == PdfStandardFont.Helvetica ||
            font == PdfStandardFont.HelveticaBold ||
            font == PdfStandardFont.HelveticaOblique ||
            font == PdfStandardFont.HelveticaBoldOblique) {
            return HelveticaGlyphWidthEmFor(value, font);
        }

        if (font == PdfStandardFont.TimesRoman ||
            font == PdfStandardFont.TimesBold ||
            font == PdfStandardFont.TimesItalic ||
            font == PdfStandardFont.TimesBoldItalic) {
            return TimesGlyphWidthEmFor(value, font);
        }

        return GlyphWidthEmFor(font);
    }

    private static double HelveticaGlyphWidthEmFor(char value, PdfStandardFont font) {
        if (font == PdfStandardFont.HelveticaBold || font == PdfStandardFont.HelveticaBoldOblique) {
            return HelveticaBoldGlyphWidthEmFor(value, font);
        }

        if (TryGetHelveticaWinAnsiWidth1000(value, bold: false, out int winAnsiWidth)) {
            return winAnsiWidth / 1000D;
        }

        if (TryGetWinAnsiLatinBaseChar(value, out char baseChar)) {
            return HelveticaGlyphWidthEmFor(baseChar, font);
        }

        int width = value switch {
            ' ' => 278,
            '!' => 278,
            '"' => 355,
            '#' => 556,
            '$' => 556,
            '%' => 889,
            '&' => 667,
            '\'' => 222,
            '(' => 333,
            ')' => 333,
            '*' => 389,
            '+' => 584,
            ',' => 278,
            '-' => 333,
            '.' => 278,
            '/' => 278,
            >= '0' and <= '9' => 556,
            ':' => 278,
            ';' => 278,
            '<' => 584,
            '=' => 584,
            '>' => 584,
            '?' => 556,
            '@' => 1015,
            'A' => 667,
            'B' => 667,
            'C' => 722,
            'D' => 722,
            'E' => 667,
            'F' => 611,
            'G' => 778,
            'H' => 722,
            'I' => 278,
            'J' => 500,
            'K' => 667,
            'L' => 556,
            'M' => 833,
            'N' => 722,
            'O' => 778,
            'P' => 667,
            'Q' => 778,
            'R' => 722,
            'S' => 667,
            'T' => 611,
            'U' => 722,
            'V' => 667,
            'W' => 944,
            'X' => 667,
            'Y' => 667,
            'Z' => 611,
            '[' => 278,
            '\\' => 278,
            ']' => 278,
            '^' => 469,
            '_' => 556,
            '`' => 222,
            'a' => 556,
            'b' => 556,
            'c' => 500,
            'd' => 556,
            'e' => 556,
            'f' => 278,
            'g' => 556,
            'h' => 556,
            'i' => 222,
            'j' => 222,
            'k' => 500,
            'l' => 222,
            'm' => 833,
            'n' => 556,
            'o' => 556,
            'p' => 556,
            'q' => 556,
            'r' => 333,
            's' => 500,
            't' => 278,
            'u' => 556,
            'v' => 500,
            'w' => 722,
            'x' => 500,
            'y' => 500,
            'z' => 500,
            '{' => 334,
            '|' => 260,
            '}' => 334,
            '~' => 584,
            _ => (int)(GlyphWidthEmFor(font) * 1000)
        };

        return width / 1000D;
    }

    private static double HelveticaBoldGlyphWidthEmFor(char value, PdfStandardFont font) {
        if (value >= ' ' && value <= '~') {
            return HelveticaBoldAsciiWidths[value - ' '] / 1000D;
        }

        if (TryGetHelveticaWinAnsiWidth1000(value, bold: true, out int winAnsiWidth)) {
            return winAnsiWidth / 1000D;
        }

        if (TryGetWinAnsiLatinBaseChar(value, out char baseChar)) {
            return HelveticaBoldGlyphWidthEmFor(baseChar, font);
        }

        return GlyphWidthEmFor(font);
    }

    private static bool TryGetHelveticaWinAnsiWidth1000(char value, bool bold, out int width) {
        width = value switch {
            '\u00A0' => 278,
            '€' => 556,
            '£' => 556,
            '¥' => 556,
            '¢' => 556,
            '©' => 737,
            '®' => 737,
            '§' => 556,
            '°' => 400,
            '±' => 584,
            '•' => 350,
            '–' => 556,
            '—' => 1000,
            '‘' => bold ? 278 : 222,
            '’' => bold ? 278 : 222,
            '‚' => bold ? 278 : 222,
            '“' => bold ? 500 : 333,
            '”' => bold ? 500 : 333,
            '„' => bold ? 500 : 333,
            '…' => 1000,
            '†' => 556,
            '‡' => 556,
            '‰' => 1000,
            '‹' => bold ? 333 : 333,
            '›' => bold ? 333 : 333,
            '™' => 1000,
            _ => -1
        };

        return width >= 0;
    }

    private static double TimesGlyphWidthEmFor(char value, PdfStandardFont font) {
        if (font == PdfStandardFont.TimesBold) {
            return TimesVariantGlyphWidthEmFor(value, font, TimesBoldAsciiWidths);
        }

        if (font == PdfStandardFont.TimesItalic) {
            return TimesVariantGlyphWidthEmFor(value, font, TimesItalicAsciiWidths);
        }

        if (font == PdfStandardFont.TimesBoldItalic) {
            return TimesVariantGlyphWidthEmFor(value, font, TimesBoldItalicAsciiWidths);
        }

        if (TryGetTimesWinAnsiWidth1000(value, font, out int winAnsiWidth)) {
            return winAnsiWidth / 1000D;
        }

        if (TryGetWinAnsiLatinBaseChar(value, out char baseChar)) {
            return TimesGlyphWidthEmFor(baseChar, font);
        }

        int width = value switch {
            ' ' => 250,
            '!' => 333,
            '"' => 408,
            '#' => 500,
            '$' => 500,
            '%' => 833,
            '&' => 778,
            '\'' => 180,
            '(' => 333,
            ')' => 333,
            '*' => 500,
            '+' => 564,
            ',' => 250,
            '-' => 333,
            '.' => 250,
            '/' => 278,
            >= '0' and <= '9' => 500,
            ':' => 278,
            ';' => 278,
            '<' => 564,
            '=' => 564,
            '>' => 564,
            '?' => 444,
            '@' => 921,
            'A' => 722,
            'B' => 667,
            'C' => 667,
            'D' => 722,
            'E' => 611,
            'F' => 556,
            'G' => 722,
            'H' => 722,
            'I' => 333,
            'J' => 389,
            'K' => 722,
            'L' => 611,
            'M' => 889,
            'N' => 722,
            'O' => 722,
            'P' => 556,
            'Q' => 722,
            'R' => 667,
            'S' => 556,
            'T' => 611,
            'U' => 722,
            'V' => 722,
            'W' => 944,
            'X' => 722,
            'Y' => 722,
            'Z' => 611,
            '[' => 333,
            '\\' => 278,
            ']' => 333,
            '^' => 469,
            '_' => 500,
            '`' => 333,
            'a' => 444,
            'b' => 500,
            'c' => 444,
            'd' => 500,
            'e' => 444,
            'f' => 333,
            'g' => 500,
            'h' => 500,
            'i' => 278,
            'j' => 278,
            'k' => 500,
            'l' => 278,
            'm' => 778,
            'n' => 500,
            'o' => 500,
            'p' => 500,
            'q' => 500,
            'r' => 333,
            's' => 389,
            't' => 278,
            'u' => 500,
            'v' => 500,
            'w' => 722,
            'x' => 500,
            'y' => 500,
            'z' => 444,
            '{' => 480,
            '|' => 200,
            '}' => 480,
            '~' => 541,
            _ => (int)(GlyphWidthEmFor(font) * 1000)
        };

        return width / 1000D;
    }

    private static double TimesVariantGlyphWidthEmFor(char value, PdfStandardFont font, int[] asciiWidths) {
        if (value >= ' ' && value <= '~') {
            return asciiWidths[value - ' '] / 1000D;
        }

        if (TryGetTimesWinAnsiWidth1000(value, font, out int winAnsiWidth)) {
            return winAnsiWidth / 1000D;
        }

        if (TryGetWinAnsiLatinBaseChar(value, out char baseChar)) {
            return TimesVariantGlyphWidthEmFor(baseChar, font, asciiWidths);
        }

        return GlyphWidthEmFor(font);
    }

    private static bool TryGetTimesWinAnsiWidth1000(char value, PdfStandardFont font, out int width) {
        bool bold = font == PdfStandardFont.TimesBold || font == PdfStandardFont.TimesBoldItalic;
        bool italic = font == PdfStandardFont.TimesItalic || font == PdfStandardFont.TimesBoldItalic;
        width = value switch {
            '\u00A0' => 250,
            '€' => 500,
            '£' => 500,
            '¥' => 500,
            '¢' => 500,
            '©' => 760,
            '®' => 760,
            '§' => 500,
            '°' => 400,
            '±' => 564,
            '•' => 350,
            '–' => 500,
            '—' => italic && !bold ? 889 : 1000,
            '‘' => 333,
            '’' => 333,
            '‚' => 333,
            '“' => bold ? 500 : italic ? 556 : 444,
            '”' => bold ? 500 : italic ? 556 : 444,
            '„' => bold ? 500 : italic ? 556 : 444,
            '…' => italic && !bold ? 889 : 1000,
            '†' => 500,
            '‡' => 500,
            '‰' => 1000,
            '‹' => 333,
            '›' => 333,
            '™' => bold ? 1000 : 980,
            _ => -1
        };

        return width >= 0;
    }

    private static bool TryGetWinAnsiLatinBaseChar(char value, out char baseChar) {
        baseChar = value switch {
            'À' or 'Á' or 'Â' or 'Ã' or 'Ä' or 'Å' => 'A',
            'à' or 'á' or 'â' or 'ã' or 'ä' or 'å' => 'a',
            'Ç' => 'C',
            'ç' => 'c',
            'È' or 'É' or 'Ê' or 'Ë' => 'E',
            'è' or 'é' or 'ê' or 'ë' => 'e',
            'Ì' or 'Í' or 'Î' or 'Ï' => 'I',
            'ì' or 'í' or 'î' or 'ï' => 'i',
            'Ñ' => 'N',
            'ñ' => 'n',
            'Ò' or 'Ó' or 'Ô' or 'Õ' or 'Ö' or 'Ø' => 'O',
            'ò' or 'ó' or 'ô' or 'õ' or 'ö' or 'ø' => 'o',
            'Š' => 'S',
            'š' => 's',
            'Ù' or 'Ú' or 'Û' or 'Ü' => 'U',
            'ù' or 'ú' or 'û' or 'ü' => 'u',
            'Ý' or 'Ÿ' => 'Y',
            'ý' or 'ÿ' => 'y',
            'Ž' => 'Z',
            'ž' => 'z',
            'Ð' => 'D',
            'ð' => 'o',
            'Þ' => 'P',
            'þ' => 'p',
            'µ' => 'u',
            _ => '\0'
        };

        return baseChar != '\0';
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
