using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const int MaximumTextLayoutLines = 100_000;
    private const double DefaultParagraphTabStopWidth = 36D;
    private static readonly char[] TokenSplitChars = new[] { ' ', '\n', '\t' };
    private static readonly char[] HardLineSplitChars = new[] { '\n' };
    private static readonly char[] SoftLineSplitChars = new[] { ' ', '\t' };
    private static readonly char[] DecimalTabAnchorChars = new[] { '.', ',' };
    private static readonly char[] LongTokenDelimiterBreakChars = new[] { '-', '.', '_', '/', '\\', ':', '|' };
    private static int[] GetValidHyphenationBreakpoints(string token, PdfOptions? options) {
        PdfTextHyphenationCallback? callback = options?.TextHyphenationCallbackSnapshot;
        if (callback == null || string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        System.Collections.Generic.IReadOnlyList<int>? points = callback(token);
        if (points == null || points.Count == 0) {
            return Array.Empty<int>();
        }

        return points
            .Where(point => IsValidTokenBreakIndex(token, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
    }

    private static bool IsValidTokenBreakIndex(string token, int index) =>
        index > 0 &&
        index < token.Length &&
        !(index > 0 && index < token.Length && char.IsHighSurrogate(token[index - 1]) && char.IsLowSurrogate(token[index]));

    private static int[] GetValidLongTokenDelimiterBreakpoints(string token) {
        if (string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        return token
            .Select((ch, index) => IsLongTokenDelimiterBreakChar(ch) ? index + 1 : -1)
            .Where(point => IsValidTokenBreakIndex(token, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
    }

    private static bool IsLongTokenDelimiterBreakChar(char value) =>
        Array.IndexOf(LongTokenDelimiterBreakChars, value) >= 0;

    private static PdfTextRun CreateStyledTextRun(string text, PdfTextRun styleTemplate, PdfStandardFont? font, string? fallbackFontFamily = null) {
        bool keepLink = !string.IsNullOrWhiteSpace(text) &&
            (styleTemplate.LinkUri != null || styleTemplate.LinkDestinationName != null);

        return new PdfTextRun(
            text,
            styleTemplate.Bold,
            styleTemplate.Underline,
            styleTemplate.Color,
            styleTemplate.Italic,
            styleTemplate.Strike,
            styleTemplate.FontSize,
            font,
            keepLink ? styleTemplate.LinkUri : null,
            keepLink ? styleTemplate.LinkContents : null,
            styleTemplate.Baseline,
            keepLink ? styleTemplate.LinkDestinationName : null,
            backgroundColor: styleTemplate.BackgroundColor,
            fontFamily: styleTemplate.FontFamily ?? fallbackFontFamily,
            underlineStyle: styleTemplate.UnderlineStyle,
            strikeStyle: styleTemplate.StrikeStyle,
            decorationColor: styleTemplate.DecorationColor)
            .WithSpacingFrom(styleTemplate).WithFeatureSettings(styleTemplate.FeatureSettings)
            .WithHorizontalOffset(styleTemplate.HorizontalOffset)
            .WithTextDirection(styleTemplate.TextDirection);
    }

    private static bool CanWriteRunWithSelectedFont(PdfTextRun run, PdfStandardFont baseFont, PdfOptions? options) {
        string text = run.Text ?? string.Empty;
        if (text.Length == 0 || IsLayoutControlRun(run)) {
            return true;
        }

        if (options != null &&
            options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace namedFace)) {
            if (options.TryGetNamedFontProgram(namedFace, out PdfTrueTypeFontProgram? namedFontProgram) &&
                namedFontProgram != null) {
                return CanWriteWithEmbeddedFont(text, namedFontProgram, options.TextShapingModeSnapshot);
            }

            if (options.TryGetNamedOpenTypeCffFontProgram(namedFace, out PdfOpenTypeCffFontProgram? namedCffFontProgram) &&
                namedCffFontProgram != null) {
                return CanWriteWithEmbeddedFont(text, namedCffFontProgram, options.TextShapingModeSnapshot);
            }
        }

        PdfStandardFont fontForRun = ResolveFontForRun(run, baseFont);
        return CanWriteTextWithSelectedFont(text, fontForRun, options);
    }

    private static bool CanWriteTextWithSelectedFont(string text, PdfStandardFont fontForRun, PdfOptions? options) {
        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(fontForRun, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return CanWriteWithEmbeddedFont(text, fontProgram, options.TextShapingModeSnapshot);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(fontForRun, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return CanWriteWithEmbeddedFont(text, cffFontProgram, options.TextShapingModeSnapshot);
        }

        return PdfWinAnsiEncoding.CanEncode(text, out _);
    }

    private static PdfStandardFont ResolveFontForRun(PdfTextRun run, PdfStandardFont baseFont) {
        PdfStandardFont runBaseFont = run.Font.HasValue ? ChooseNormal(run.Font.Value) : baseFont;
        return (run.Bold && run.Italic)
            ? ChooseBoldItalic(runBaseFont)
            : run.Bold
                ? ChooseBold(runBaseFont)
                : run.Italic
                    ? ChooseItalic(runBaseFont)
                    : runBaseFont;
    }

    private static bool CanWriteWithEmbeddedFont(string text, PdfTrueTypeFontProgram fontProgram, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        int index = 0;
        while (index < text.Length) {
            int scalarStart = index;
            if (shapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                return false;
            }

            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfTrueTypeFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (shapingMode == PdfTextShapingMode.LatinLigatures &&
            OfficeTextLigatures.TryGetLatinPresentationForm(text, index, out int ligatureScalar, out length) &&
            fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
            ligatureGlyphId > 0) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool CanWriteWithEmbeddedFont(string text, PdfOpenTypeCffFontProgram fontProgram, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        int index = 0;
        while (index < text.Length) {
            int scalarStart = index;
            if (shapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                return false;
            }

            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfOpenTypeCffFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (shapingMode == PdfTextShapingMode.LatinLigatures &&
            OfficeTextLigatures.TryGetLatinPresentationForm(text, index, out int ligatureScalar, out length) &&
            fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
            ligatureGlyphId > 0) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool IsLayoutControlRun(PdfTextRun run) =>
        string.Equals(run.Text, "\n", StringComparison.Ordinal) ||
        string.Equals(run.Text, "\t", StringComparison.Ordinal);

    private static PdfAlign ResolveRichLineAlignment(PdfAlign fallback, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments, int lineIndex) =>
        lineAlignments != null && lineIndex >= 0 && lineIndex < lineAlignments.Count && lineAlignments[lineIndex].HasValue
            ? lineAlignments[lineIndex]!.Value
            : fallback;

    private static double ResolveRichLineWidth(double fallback, double? firstLineWidthOverride, System.Collections.Generic.IReadOnlyList<double>? lineWidths, int lineIndex) =>
        lineWidths != null && lineIndex >= 0 && lineIndex < lineWidths.Count
            ? lineWidths[lineIndex]
            : lineIndex == 0 ? firstLineWidthOverride ?? fallback : fallback;

    private static double ResolveRichLineXOrigin(double fallback, double? firstLineXOverride, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets, int lineIndex) =>
        lineXOffsets != null && lineIndex >= 0 && lineIndex < lineXOffsets.Count
            ? fallback + lineXOffsets[lineIndex]
            : lineIndex == 0 ? firstLineXOverride ?? fallback : fallback;

    private static double GetRichSegmentWidth(RichSeg segment) =>
        segment.MeasuredWidth;
}
