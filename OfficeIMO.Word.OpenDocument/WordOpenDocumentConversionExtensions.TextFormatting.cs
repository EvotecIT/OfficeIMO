using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

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
}
