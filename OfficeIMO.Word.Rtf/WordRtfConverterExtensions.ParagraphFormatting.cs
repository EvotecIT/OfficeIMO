using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyParagraphFormatting(WordParagraph source, RtfParagraph destination, RtfDocument document) {
        CopyParagraphStyleAndNumbering(source, destination, document);
        CopyTabStops(source, destination);
        CopyParagraphLayout(source, destination);

        string shading = source.ShadingFillValue;
        if (!string.IsNullOrWhiteSpace(shading) &&
            !string.Equals(shading, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(shading, out byte red, out byte green, out byte blue)) {
            destination.BackgroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        string? foreground = source.ParagraphPropertiesNode().WChild("shd").WAttr("color");
        if (!string.IsNullOrWhiteSpace(foreground) &&
            !string.Equals(foreground, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(foreground!, out red, out green, out blue)) {
            destination.ShadingForegroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        CopyParagraphShadingPattern(source.ShadingPattern, destination);
        CopyParagraphBorder(source.Borders.TopStyle, source.Borders.TopSize, source.Borders.TopColorValue, destination.TopBorder, document);
        CopyParagraphBorder(source.Borders.LeftStyle, source.Borders.LeftSize, source.Borders.LeftColorValue, destination.LeftBorder, document);
        CopyParagraphBorder(source.Borders.BottomStyle, source.Borders.BottomSize, source.Borders.BottomColorValue, destination.BottomBorder, document);
        CopyParagraphBorder(source.Borders.RightStyle, source.Borders.RightSize, source.Borders.RightColorValue, destination.RightBorder, document);
    }

    private static void CopyParagraphBorder(WordBorderStyle? style, OfficeLength? width, string? colorHex, RtfParagraphBorder destination, RtfDocument document) {
        destination.DirectStyle = style.HasValue ? ToRtfParagraphBorderStyle(style) : null;
        if (width.HasValue) {
            destination.Width = checked((int)Math.Round(width.Value.ToPoints() * 20D, MidpointRounding.AwayFromZero));
        }

        if (!string.IsNullOrWhiteSpace(colorHex) &&
            !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(colorHex!, out byte red, out byte green, out byte blue)) {
            destination.ColorIndex = GetOrAddColor(document, red, green, blue);
        }
    }

    private static void ApplyParagraphFormatting(WordParagraph destination, RtfParagraph source, RtfDocument document, RtfWordNumberingPlan numbering) {
        int? projectedNumberId = numbering.GetNumberId(source);
        source = document.GetParagraphFormatting(source);
        ApplyParagraphStyleAndNumbering(destination, source, document, projectedNumberId);
        // The common formatting path receives effective style tabs, including
        // paragraphs in tables, notes and running stories.
        ApplyTabStops(destination, source, document);
        ApplyParagraphLayout(destination, source);

        if (source.BackgroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.BackgroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                destination.ShadingFillValue = color!;
            }
        }

        OfficeXmlNode shading;
        if (source.ShadingForegroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ShadingForegroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                shading = GetOrCreateParagraphShading(destination);
                shading.SetAttribute(WordXml.W, "color", color!);
            }
        }

        WordShadingPattern? pattern = ToWordShadingPattern(source);
        if (pattern.HasValue) {
            shading = GetOrCreateParagraphShading(destination);
            shading.SetAttribute(WordXml.W, "val", pattern.Value.ToOpenXmlToken());
        }

        ApplyParagraphBorder(source.TopBorder, style => destination.Borders.TopStyle = style, width => destination.Borders.TopSize = width.HasValue ? OfficeLength.Points(width.Value / 8D) : null, color => destination.Borders.TopColorValue = color, document);
        ApplyParagraphBorder(source.LeftBorder, style => destination.Borders.LeftStyle = style, width => destination.Borders.LeftSize = width.HasValue ? OfficeLength.Points(width.Value / 8D) : null, color => destination.Borders.LeftColorValue = color, document);
        ApplyParagraphBorder(source.BottomBorder, style => destination.Borders.BottomStyle = style, width => destination.Borders.BottomSize = width.HasValue ? OfficeLength.Points(width.Value / 8D) : null, color => destination.Borders.BottomColorValue = color, document);
        ApplyParagraphBorder(source.RightBorder, style => destination.Borders.RightStyle = style, width => destination.Borders.RightSize = width.HasValue ? OfficeLength.Points(width.Value / 8D) : null, color => destination.Borders.RightColorValue = color, document);
    }

    private static void CopyParagraphShadingPattern(WordShadingPattern? source, RtfParagraph destination) {
        if (!source.HasValue) {
            return;
        }

        int? shading = ToRtfShadingPercent(source.Value);
        if (shading.HasValue) {
            destination.ShadingPatternPercent = shading;
        }

        destination.ShadingPattern = ToRtfShadingPattern(source.Value);
    }

    private static OfficeXmlNode GetOrCreateParagraphShading(WordParagraph paragraph) =>
        GetOrAddParagraphProperties(paragraph).GetOrAddChild(WordXml.W, "shd", WPPrView.Order);

    private static int? ToRtfShadingPercent(WordShadingPattern value) {
        if (value == WordShadingPattern.Solid) return 10000;
        if (value == WordShadingPattern.Percent5) return 500;
        if (value == WordShadingPattern.Percent10) return 1000;
        if (value == WordShadingPattern.Percent12) return 1250;
        if (value == WordShadingPattern.Percent15) return 1500;
        if (value == WordShadingPattern.Percent20) return 2000;
        if (value == WordShadingPattern.Percent25) return 2500;
        if (value == WordShadingPattern.Percent30) return 3000;
        if (value == WordShadingPattern.Percent35) return 3500;
        if (value == WordShadingPattern.Percent37) return 3750;
        if (value == WordShadingPattern.Percent40) return 4000;
        if (value == WordShadingPattern.Percent45) return 4500;
        if (value == WordShadingPattern.Percent50) return 5000;
        if (value == WordShadingPattern.Percent55) return 5500;
        if (value == WordShadingPattern.Percent60) return 6000;
        if (value == WordShadingPattern.Percent62) return 6250;
        if (value == WordShadingPattern.Percent65) return 6500;
        if (value == WordShadingPattern.Percent70) return 7000;
        if (value == WordShadingPattern.Percent75) return 7500;
        if (value == WordShadingPattern.Percent80) return 8000;
        if (value == WordShadingPattern.Percent85) return 8500;
        if (value == WordShadingPattern.Percent87) return 8750;
        if (value == WordShadingPattern.Percent90) return 9000;
        if (value == WordShadingPattern.Percent95) return 9500;
        return null;
    }

    private static RtfShadingPattern ToRtfShadingPattern(WordShadingPattern value) {
        if (value == WordShadingPattern.ThinHorizontalStripe) return RtfShadingPattern.Horizontal;
        if (value == WordShadingPattern.ThinVerticalStripe) return RtfShadingPattern.Vertical;
        if (value == WordShadingPattern.ThinDiagonalStripe) return RtfShadingPattern.ForwardDiagonal;
        if (value == WordShadingPattern.ThinReverseDiagonalStripe) return RtfShadingPattern.BackwardDiagonal;
        if (value == WordShadingPattern.ThinHorizontalCross) return RtfShadingPattern.Cross;
        if (value == WordShadingPattern.ThinDiagonalCross) return RtfShadingPattern.DiagonalCross;
        if (value == WordShadingPattern.HorizontalStripe) return RtfShadingPattern.DarkHorizontal;
        if (value == WordShadingPattern.VerticalStripe) return RtfShadingPattern.DarkVertical;
        if (value == WordShadingPattern.DiagonalStripe) return RtfShadingPattern.DarkForwardDiagonal;
        if (value == WordShadingPattern.ReverseDiagonalStripe) return RtfShadingPattern.DarkBackwardDiagonal;
        if (value == WordShadingPattern.HorizontalCross) return RtfShadingPattern.DarkCross;
        if (value == WordShadingPattern.DiagonalCross) return RtfShadingPattern.DarkDiagonalCross;
        return RtfShadingPattern.None;
    }

    private static WordShadingPattern? ToWordShadingPattern(RtfParagraph source) {
        WordShadingPattern? percentPattern = ToWordShadingPercent(source.ShadingPatternPercent);
        if (percentPattern.HasValue) {
            return percentPattern;
        }

        return ToWordShadingPattern(source.ShadingPattern);
    }

    private static WordShadingPattern? ToWordShadingPercent(int? value) {
        switch (value) {
            case 500:
                return WordShadingPattern.Percent5;
            case 1000:
                return WordShadingPattern.Percent10;
            case 1250:
                return WordShadingPattern.Percent12;
            case 1500:
                return WordShadingPattern.Percent15;
            case 2000:
                return WordShadingPattern.Percent20;
            case 2500:
                return WordShadingPattern.Percent25;
            case 3000:
                return WordShadingPattern.Percent30;
            case 3500:
                return WordShadingPattern.Percent35;
            case 3750:
                return WordShadingPattern.Percent37;
            case 4000:
                return WordShadingPattern.Percent40;
            case 4500:
                return WordShadingPattern.Percent45;
            case 5000:
                return WordShadingPattern.Percent50;
            case 5500:
                return WordShadingPattern.Percent55;
            case 6000:
                return WordShadingPattern.Percent60;
            case 6250:
                return WordShadingPattern.Percent62;
            case 6500:
                return WordShadingPattern.Percent65;
            case 7000:
                return WordShadingPattern.Percent70;
            case 7500:
                return WordShadingPattern.Percent75;
            case 8000:
                return WordShadingPattern.Percent80;
            case 8500:
                return WordShadingPattern.Percent85;
            case 8750:
                return WordShadingPattern.Percent87;
            case 9000:
                return WordShadingPattern.Percent90;
            case 9500:
                return WordShadingPattern.Percent95;
            case 10000:
                return WordShadingPattern.Solid;
            default:
                return null;
        }
    }

    private static WordShadingPattern? ToWordShadingPattern(RtfShadingPattern pattern) {
        switch (pattern) {
            case RtfShadingPattern.Horizontal:
                return WordShadingPattern.ThinHorizontalStripe;
            case RtfShadingPattern.Vertical:
                return WordShadingPattern.ThinVerticalStripe;
            case RtfShadingPattern.ForwardDiagonal:
                return WordShadingPattern.ThinDiagonalStripe;
            case RtfShadingPattern.BackwardDiagonal:
                return WordShadingPattern.ThinReverseDiagonalStripe;
            case RtfShadingPattern.Cross:
                return WordShadingPattern.ThinHorizontalCross;
            case RtfShadingPattern.DiagonalCross:
                return WordShadingPattern.ThinDiagonalCross;
            case RtfShadingPattern.DarkHorizontal:
                return WordShadingPattern.HorizontalStripe;
            case RtfShadingPattern.DarkVertical:
                return WordShadingPattern.VerticalStripe;
            case RtfShadingPattern.DarkForwardDiagonal:
                return WordShadingPattern.DiagonalStripe;
            case RtfShadingPattern.DarkBackwardDiagonal:
                return WordShadingPattern.ReverseDiagonalStripe;
            case RtfShadingPattern.DarkCross:
                return WordShadingPattern.HorizontalCross;
            case RtfShadingPattern.DarkDiagonalCross:
                return WordShadingPattern.DiagonalCross;
            default:
                return null;
        }
    }

    private static void ApplyParagraphBorder(RtfParagraphBorder source, Action<WordBorderStyle?> setStyle, Action<uint?> setWidth, Action<string?> setColor, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        setStyle(ToWordParagraphBorderStyle(source.Style));
        if (source.Width.HasValue && source.Width.Value >= 0) {
            setWidth(ToWordBorderSize(source.Width.Value));
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                setColor(color);
            }
        }
    }

    private static RtfParagraphBorderStyle ToRtfParagraphBorderStyle(WordBorderStyle? value) {
        if (value == WordBorderStyle.Double) return RtfParagraphBorderStyle.Double;
        if (value == WordBorderStyle.Dotted) return RtfParagraphBorderStyle.Dotted;
        if (value == WordBorderStyle.Dashed) return RtfParagraphBorderStyle.Dashed;
        if (value == WordBorderStyle.Nil || value == WordBorderStyle.None) return RtfParagraphBorderStyle.None;
        if (value == WordBorderStyle.Single) return RtfParagraphBorderStyle.Single;
        return RtfParagraphBorderStyle.None;
    }

    private static WordBorderStyle? ToWordParagraphBorderStyle(RtfParagraphBorderStyle value) {
        switch (value) {
            case RtfParagraphBorderStyle.Double:
                return WordBorderStyle.Double;
            case RtfParagraphBorderStyle.Dotted:
                return WordBorderStyle.Dotted;
            case RtfParagraphBorderStyle.Dashed:
                return WordBorderStyle.Dashed;
            case RtfParagraphBorderStyle.Single:
                return WordBorderStyle.Single;
            default:
                return WordBorderStyle.Nil;
        }
    }
}
