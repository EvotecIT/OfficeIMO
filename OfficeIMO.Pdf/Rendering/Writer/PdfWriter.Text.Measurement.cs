using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void MarkRichLineTextSeparator(System.Collections.Generic.IList<RichSeg> line) {
        if (line.Count == 0) {
            return;
        }

        int lastIndex = line.Count - 1;
        line[lastIndex] = line[lastIndex].WithEndsWithTextSeparator();
    }

    private static void MarkRichLineHardBreak(System.Collections.Generic.IList<RichSeg> line) {
        if (line.Count == 0) {
            return;
        }

        int lastIndex = line.Count - 1;
        line[lastIndex] = line[lastIndex].WithEndsWithHardBreak();
    }

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize, PdfOptions? options = null) =>
        EstimateSimpleTextWidthForOptions(text, font, fontSize, options);

    private static double MeasureRichText(string text, PdfStandardFont font, PdfNamedFontFace? namedFont, double fontSize, PdfOptions? options = null) =>
        EstimateSimpleTextWidthForOptions(text, font, namedFont, fontSize, options);

    private static double EffectiveRichFontSize(double fontSize, PdfTextBaseline baseline) =>
        baseline == PdfTextBaseline.Normal ? fontSize : fontSize * 0.65;

    private static double TextRiseForBaseline(double fontSize, PdfTextBaseline baseline) => baseline switch {
        PdfTextBaseline.Superscript => fontSize * 0.35,
        PdfTextBaseline.Subscript => -fontSize * 0.18,
        _ => 0
    };

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize, PdfTextBaseline baseline, PdfOptions? options = null) =>
        EstimateSimpleTextWidthForOptions(text, font, EffectiveRichFontSize(fontSize, baseline), options);

    private static double MeasureRichText(string text, PdfStandardFont font, PdfNamedFontFace? namedFont, double fontSize, PdfTextBaseline baseline, PdfOptions? options = null, OfficeTextFeatureSettings? featureSettings = null, double horizontalTextScaling = 100D, double characterSpacing = 0D, OfficeTextDirection textDirection = OfficeTextDirection.Auto) {
        double effectiveFontSize = EffectiveRichFontSize(fontSize, baseline);
        if (horizontalTextScaling == 100D && characterSpacing == 0D) {
            return EstimateSimpleTextWidthForOptions(text, font, namedFont, effectiveFontSize, options, featureSettings);
        }
        if (ContainsLineBreak(text)) {
            double maximum = 0D;
            foreach (string line in text.Split(LayoutLineSeparators, StringSplitOptions.None)) {
                maximum = Math.Max(maximum, MeasureRichText(line, font, namedFont, fontSize, baseline, options,
                    featureSettings, horizontalTextScaling, characterSpacing, textDirection));
            }
            return maximum;
        }

        PdfTextShowCommand command = EncodeTextShowCommand(text, font, namedFont, options, featureSettings, textDirection);
        double naturalWidth = command.AdvanceWidth1000.GetValueOrDefault() * effectiveFontSize / 1000D;
        double advance = naturalWidth * horizontalTextScaling / 100D + command.GlyphCount * characterSpacing;
        if (advance < 0D || double.IsNaN(advance) || double.IsInfinity(advance)) {
            throw new InvalidOperationException("The requested glyph width and character spacing produce an invalid text advance.");
        }
        return advance;
    }

    private static double MeasureRichLineWidth(System.Collections.Generic.IReadOnlyList<RichSeg> line, PdfOptions? options = null) {
        double width = 0D;
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (segment.LeadingSpace) {
                width += segment.LeadingAdvance > 0
                    ? segment.LeadingAdvance
                    : MeasureRichText(" ", segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline, options, segment.FeatureSettings, segment.HorizontalTextScaling, segment.CharacterSpacing);
            }

            width += GetRichSegmentWidth(segment);
        }

        return width;
    }

    private static double CalculateDefaultTabAdvance(double lineWidth, double spaceWidth, double tabStopWidth = DefaultParagraphTabStopWidth) {
        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            return spaceWidth;
        }

        double nextStop = (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        return Math.Max(spaceWidth, nextStop - lineWidth);
    }

    private static double CalculateTabAdvance(double lineWidth, double followingTextWidth, double spaceWidth, PdfTabAlignment alignment, double tabStopWidth = DefaultParagraphTabStopWidth, string followingText = "", PdfStandardFont followingFont = PdfStandardFont.Helvetica, double fontSize = 12D, PdfTextBaseline baseline = PdfTextBaseline.Normal, PdfOptions? options = null, double? maxWidth = null, PdfTabStop? explicitTabStop = null, double lineOriginOffset = 0D, PdfNamedFontFace? followingNamedFont = null, OfficeTextFeatureSettings? featureSettings = null, double horizontalTextScaling = 100D, double characterSpacing = 0D) {
        if (explicitTabStop == null && alignment == PdfTabAlignment.Left) {
            return CalculateDefaultTabAdvance(lineWidth, spaceWidth, tabStopWidth);
        }

        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            followingTextWidth < 0 || double.IsNaN(followingTextWidth) || double.IsInfinity(followingTextWidth) ||
            double.IsNaN(lineOriginOffset) || double.IsInfinity(lineOriginOffset)) {
            return spaceWidth;
        }

        if (explicitTabStop == null &&
            (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth))) {
            return spaceWidth;
        }

        double? boundedMaxWidth = maxWidth.HasValue &&
            maxWidth.Value > 0 &&
            !double.IsNaN(maxWidth.Value) &&
            !double.IsInfinity(maxWidth.Value)
                ? maxWidth.Value
                : null;
        if (explicitTabStop != null) {
            alignment = explicitTabStop.Alignment;
        }

        double anchorWidth = alignment switch {
            PdfTabAlignment.Center => followingTextWidth / 2D,
            PdfTabAlignment.Right => followingTextWidth,
            PdfTabAlignment.DecimalSeparator => MeasureDecimalAnchorWidth(followingText, followingFont, fontSize, baseline, options, followingNamedFont, featureSettings, horizontalTextScaling, characterSpacing),
            _ => 0D
        };
        double nextStop = explicitTabStop?.Position - lineOriginOffset ?? (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        if (boundedMaxWidth.HasValue) {
            nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
        }

        double advance = nextStop - anchorWidth - lineWidth;
        if (explicitTabStop != null) {
            return Math.Max(0D, advance);
        }

        if (advance < spaceWidth) {
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }

            double stopsToAdd = Math.Ceiling((spaceWidth - advance) / tabStopWidth);
            nextStop += Math.Max(1D, stopsToAdd) * tabStopWidth;
            if (boundedMaxWidth.HasValue) {
                nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
            }

            advance = nextStop - anchorWidth - lineWidth;
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }
        }

        return Math.Max(spaceWidth, advance);
    }
}
