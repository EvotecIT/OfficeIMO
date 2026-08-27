using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static int ApplyOdtHyperlinkFormatting(OdtHyperlink source, OdtParagraph paragraph,
        WordParagraph target, ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = (source.Underline ?? paragraph.Underline) == true
            ? WordUnderlineStyle.Single
            : (WordUnderlineStyle?)null;
        target.Strike = (source.StrikeThrough ?? paragraph.StrikeThrough) == true;
        ApplyOdfRunSemantics(
            source.Underline ?? paragraph.Underline,
            source.UnderlineStyle ?? paragraph.UnderlineStyle,
            source.UnderlineType ?? paragraph.UnderlineType,
            source.StrikeThrough ?? paragraph.StrikeThrough,
            source.LineThroughStyle ?? paragraph.LineThroughStyle,
            source.LineThroughType ?? paragraph.LineThroughType,
            source.TextPosition ?? paragraph.TextPosition,
            source.TextTransform ?? paragraph.TextTransform,
            source.SmallCaps ?? paragraph.SmallCaps,
            target);
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdtFontSize(fontSize, target);
        string? selectedFontFamily = SelectOdfFontFamily(source.FontFamily ?? paragraph.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (selectedFontFamily != null) target.FontFamily = selectedFontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.ColorHex = color.Value.ToString();
        ApplyOdfTextBackground(source.BackgroundColor ?? paragraph.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtParagraphTextFormatting(OdtParagraph source, WordParagraph target,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        target.Bold = source.Bold == true;
        target.Italic = source.Italic == true;
        target.Underline = source.Underline == true ? WordUnderlineStyle.Single : (WordUnderlineStyle?)null;
        target.Strike = source.StrikeThrough == true;
        ApplyOdfRunSemantics(
            source.Underline,
            source.UnderlineStyle,
            source.UnderlineType,
            source.StrikeThrough,
            source.LineThroughStyle,
            source.LineThroughType,
            source.TextPosition,
            source.TextTransform,
            source.SmallCaps,
            target);
        int unsupported = ApplyOdtFontSize(source.FontSize, target);
        string? fontFamily = SelectOdfFontFamily(source.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (fontFamily != null) target.FontFamily = fontFamily;
        if (source.Color.HasValue) target.ColorHex = source.Color.Value.ToString();
        ApplyOdfTextBackground(source.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtSpanFormatting(OdtSpan source, OdtParagraph paragraph,
        WordParagraph target, ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = (source.Underline ?? paragraph.Underline) == true
            ? WordUnderlineStyle.Single
            : (WordUnderlineStyle?)null;
        target.Strike = (source.StrikeThrough ?? paragraph.StrikeThrough) == true;
        ApplyOdfRunSemantics(
            source.Underline ?? paragraph.Underline,
            source.UnderlineStyle ?? paragraph.UnderlineStyle,
            source.UnderlineType ?? paragraph.UnderlineType,
            source.StrikeThrough ?? paragraph.StrikeThrough,
            source.LineThroughStyle ?? paragraph.LineThroughStyle,
            source.LineThroughType ?? paragraph.LineThroughType,
            source.TextPosition ?? paragraph.TextPosition,
            source.TextTransform ?? paragraph.TextTransform,
            source.SmallCaps ?? paragraph.SmallCaps,
            target);
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdtFontSize(fontSize, target);
        string? selectedFontFamily = SelectOdfFontFamily(source.FontFamily ?? paragraph.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (selectedFontFamily != null) target.FontFamily = selectedFontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.ColorHex = color.Value.ToString();
        ApplyOdfTextBackground(source.BackgroundColor ?? paragraph.TextBackgroundColor, target);
        return unsupported;
    }

    private static int ApplyOdtFontSize(OdfLength? fontSize, WordParagraph target) {
        if (!fontSize.HasValue) return 0;
        if (!fontSize.Value.TryToPoints(out double points)) return 1;
        double halfPoints = points * 2D;
        double roundedHalfPoints = Math.Round(halfPoints, MidpointRounding.AwayFromZero);
        if (Math.Abs(halfPoints - roundedHalfPoints) > 0.000000001D) return 1;
        target.FontSizePoints = roundedHalfPoints / 2D;
        return 0;
    }

    private static string? SelectOdfFontFamily(string? value,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (!OdfFontFamilySyntax.TryParse(value, out OdfFontFamilySyntax? syntax)) {
            unsupportedFontFamilies++;
            return null;
        }
        if (syntax!.HasFallbacks) approximatedFontFamilyLists++;
        return syntax.PrimaryFamily;
    }

    private static void ApplyOdfTextBackground(OdfColor? source, WordParagraph target) {
        if (!source.HasValue) return;
        if (TryMapOdfHighlight(source.Value, out WordHighlightColor highlight)) target.Highlight = highlight;
        else target.RunShadingFillColorHex = source.Value.ToString();
    }

    private static void ApplyOdfRunSemantics(
        bool? underline,
        OdfTextDecorationStyle? underlineStyle,
        OdfTextDecorationType? underlineType,
        bool? strike,
        OdfTextDecorationStyle? strikeStyle,
        OdfTextDecorationType? strikeType,
        OdfTextPosition? position,
        OdfTextTransform? transform,
        bool? smallCaps,
        WordParagraph target) {
        target.Underline = MapOdfUnderline(underline, underlineStyle, underlineType);
        bool hasStrike = strike == true && strikeStyle != OdfTextDecorationStyle.None && strikeType != OdfTextDecorationType.None;
        target.DoubleStrike = hasStrike && strikeType == OdfTextDecorationType.Double;
        target.Strike = hasStrike && !target.DoubleStrike;
        target.VerticalTextAlignment = position switch {
            OdfTextPosition.Superscript => WordVerticalTextPosition.Superscript,
            OdfTextPosition.Subscript => WordVerticalTextPosition.Subscript,
            OdfTextPosition.Normal => WordVerticalTextPosition.Baseline,
            _ => null
        };
        target.CapsStyle = transform == OdfTextTransform.Uppercase
            ? WordCapsStyle.Caps
            : smallCaps == true
                ? WordCapsStyle.SmallCaps
                : WordCapsStyle.None;
        if (transform == OdfTextTransform.Lowercase) target.TransformTextCase(OfficeTextCase.Lowercase);
        else if (transform == OdfTextTransform.Capitalize) target.TransformTextCase(OfficeTextCase.TitleCase);
    }

    private static WordUnderlineStyle? MapOdfUnderline(bool? enabled, OdfTextDecorationStyle? style, OdfTextDecorationType? type) {
        if (enabled != true || style == OdfTextDecorationStyle.None || type == OdfTextDecorationType.None) return null;
        if (type == OdfTextDecorationType.Double) {
            return style == OdfTextDecorationStyle.Wave ? WordUnderlineStyle.WavyDouble : WordUnderlineStyle.Double;
        }

        return style switch {
            OdfTextDecorationStyle.Dotted => WordUnderlineStyle.Dotted,
            OdfTextDecorationStyle.Dash => WordUnderlineStyle.Dash,
            OdfTextDecorationStyle.LongDash => WordUnderlineStyle.DashLong,
            OdfTextDecorationStyle.DotDash => WordUnderlineStyle.DotDash,
            OdfTextDecorationStyle.DotDotDash => WordUnderlineStyle.DotDotDash,
            OdfTextDecorationStyle.Wave => WordUnderlineStyle.Wave,
            _ => WordUnderlineStyle.Single
        };
    }
}
