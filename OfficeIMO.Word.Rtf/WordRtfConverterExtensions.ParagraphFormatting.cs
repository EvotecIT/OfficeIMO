using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyParagraphFormatting(WordParagraph source, RtfParagraph destination, RtfDocument document) {
        CopyParagraphStyleAndNumbering(source, destination, document);
        CopyParagraphLayout(source, destination);

        string shading = source.ShadingFillColorHex;
        if (!string.IsNullOrWhiteSpace(shading) &&
            !string.Equals(shading, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(shading, out byte red, out byte green, out byte blue)) {
            destination.BackgroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        string? foreground = source._paragraphProperties?.Shading?.Color?.Value;
        if (!string.IsNullOrWhiteSpace(foreground) &&
            !string.Equals(foreground, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(foreground!, out red, out green, out blue)) {
            destination.ShadingForegroundColorIndex = GetOrAddColor(document, red, green, blue);
        }

        CopyParagraphShadingPattern(source.ShadingPattern, destination);
        CopyParagraphBorder(source.Borders.TopStyle, source.Borders.TopSize?.Value, source.Borders.TopColorHex, destination.TopBorder, document);
        CopyParagraphBorder(source.Borders.LeftStyle, source.Borders.LeftSize?.Value, source.Borders.LeftColorHex, destination.LeftBorder, document);
        CopyParagraphBorder(source.Borders.BottomStyle, source.Borders.BottomSize?.Value, source.Borders.BottomColorHex, destination.BottomBorder, document);
        CopyParagraphBorder(source.Borders.RightStyle, source.Borders.RightSize?.Value, source.Borders.RightColorHex, destination.RightBorder, document);
    }

    private static void CopyParagraphBorder(BorderValues? style, uint? width, string? colorHex, RtfParagraphBorder destination, RtfDocument document) {
        destination.Style = ToRtfParagraphBorderStyle(style);
        if (width.HasValue) {
            destination.Width = checked((int)width.Value);
        }

        if (!string.IsNullOrWhiteSpace(colorHex) &&
            !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(colorHex!, out byte red, out byte green, out byte blue)) {
            destination.ColorIndex = GetOrAddColor(document, red, green, blue);
        }
    }

    private static void ApplyParagraphFormatting(WordParagraph destination, RtfParagraph source, RtfDocument document) {
        ApplyParagraphStyleAndNumbering(destination, source, document);
        ApplyParagraphLayout(destination, source);

        if (source.BackgroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.BackgroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                destination.ShadingFillColorHex = color!;
            }
        }

        Shading? shading = null;
        if (source.ShadingForegroundColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ShadingForegroundColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                shading = GetOrCreateParagraphShading(destination);
                shading.Color = color!;
            }
        }

        ShadingPatternValues? pattern = ToWordShadingPattern(source);
        if (pattern.HasValue) {
            shading = GetOrCreateParagraphShading(destination);
            shading.Val = pattern.Value;
        }

        ApplyParagraphBorder(source.TopBorder, style => destination.Borders.TopStyle = style, width => destination.Borders.TopSize = width, color => destination.Borders.TopColorHex = color, document);
        ApplyParagraphBorder(source.LeftBorder, style => destination.Borders.LeftStyle = style, width => destination.Borders.LeftSize = width, color => destination.Borders.LeftColorHex = color, document);
        ApplyParagraphBorder(source.BottomBorder, style => destination.Borders.BottomStyle = style, width => destination.Borders.BottomSize = width, color => destination.Borders.BottomColorHex = color, document);
        ApplyParagraphBorder(source.RightBorder, style => destination.Borders.RightStyle = style, width => destination.Borders.RightSize = width, color => destination.Borders.RightColorHex = color, document);
    }

    private static void CopyParagraphShadingPattern(ShadingPatternValues? source, RtfParagraph destination) {
        if (!source.HasValue) {
            return;
        }

        int? shading = ToRtfShadingPercent(source.Value);
        if (shading.HasValue) {
            destination.ShadingPatternPercent = shading;
        }

        destination.ShadingPattern = ToRtfShadingPattern(source.Value);
    }

    private static Shading GetOrCreateParagraphShading(WordParagraph paragraph) {
        ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        properties.Shading ??= new Shading();
        return properties.Shading;
    }

    private static int? ToRtfShadingPercent(ShadingPatternValues value) {
        if (value == ShadingPatternValues.Solid) return 10000;
        if (value == ShadingPatternValues.Percent5) return 500;
        if (value == ShadingPatternValues.Percent10) return 1000;
        if (value == ShadingPatternValues.Percent12) return 1250;
        if (value == ShadingPatternValues.Percent15) return 1500;
        if (value == ShadingPatternValues.Percent20) return 2000;
        if (value == ShadingPatternValues.Percent25) return 2500;
        if (value == ShadingPatternValues.Percent30) return 3000;
        if (value == ShadingPatternValues.Percent35) return 3500;
        if (value == ShadingPatternValues.Percent37) return 3750;
        if (value == ShadingPatternValues.Percent40) return 4000;
        if (value == ShadingPatternValues.Percent45) return 4500;
        if (value == ShadingPatternValues.Percent50) return 5000;
        if (value == ShadingPatternValues.Percent55) return 5500;
        if (value == ShadingPatternValues.Percent60) return 6000;
        if (value == ShadingPatternValues.Percent62) return 6250;
        if (value == ShadingPatternValues.Percent65) return 6500;
        if (value == ShadingPatternValues.Percent70) return 7000;
        if (value == ShadingPatternValues.Percent75) return 7500;
        if (value == ShadingPatternValues.Percent80) return 8000;
        if (value == ShadingPatternValues.Percent85) return 8500;
        if (value == ShadingPatternValues.Percent87) return 8750;
        if (value == ShadingPatternValues.Percent90) return 9000;
        if (value == ShadingPatternValues.Percent95) return 9500;
        return null;
    }

    private static RtfShadingPattern ToRtfShadingPattern(ShadingPatternValues value) {
        if (value == ShadingPatternValues.ThinHorizontalStripe) return RtfShadingPattern.Horizontal;
        if (value == ShadingPatternValues.ThinVerticalStripe) return RtfShadingPattern.Vertical;
        if (value == ShadingPatternValues.ThinDiagonalStripe) return RtfShadingPattern.ForwardDiagonal;
        if (value == ShadingPatternValues.ThinReverseDiagonalStripe) return RtfShadingPattern.BackwardDiagonal;
        if (value == ShadingPatternValues.ThinHorizontalCross) return RtfShadingPattern.Cross;
        if (value == ShadingPatternValues.ThinDiagonalCross) return RtfShadingPattern.DiagonalCross;
        if (value == ShadingPatternValues.HorizontalStripe) return RtfShadingPattern.DarkHorizontal;
        if (value == ShadingPatternValues.VerticalStripe) return RtfShadingPattern.DarkVertical;
        if (value == ShadingPatternValues.DiagonalStripe) return RtfShadingPattern.DarkForwardDiagonal;
        if (value == ShadingPatternValues.ReverseDiagonalStripe) return RtfShadingPattern.DarkBackwardDiagonal;
        if (value == ShadingPatternValues.HorizontalCross) return RtfShadingPattern.DarkCross;
        if (value == ShadingPatternValues.DiagonalCross) return RtfShadingPattern.DarkDiagonalCross;
        return RtfShadingPattern.None;
    }

    private static ShadingPatternValues? ToWordShadingPattern(RtfParagraph source) {
        ShadingPatternValues? percentPattern = ToWordShadingPercent(source.ShadingPatternPercent);
        if (percentPattern.HasValue) {
            return percentPattern;
        }

        return ToWordShadingPattern(source.ShadingPattern);
    }

    private static ShadingPatternValues? ToWordShadingPercent(int? value) {
        switch (value) {
            case 500:
                return ShadingPatternValues.Percent5;
            case 1000:
                return ShadingPatternValues.Percent10;
            case 1250:
                return ShadingPatternValues.Percent12;
            case 1500:
                return ShadingPatternValues.Percent15;
            case 2000:
                return ShadingPatternValues.Percent20;
            case 2500:
                return ShadingPatternValues.Percent25;
            case 3000:
                return ShadingPatternValues.Percent30;
            case 3500:
                return ShadingPatternValues.Percent35;
            case 3750:
                return ShadingPatternValues.Percent37;
            case 4000:
                return ShadingPatternValues.Percent40;
            case 4500:
                return ShadingPatternValues.Percent45;
            case 5000:
                return ShadingPatternValues.Percent50;
            case 5500:
                return ShadingPatternValues.Percent55;
            case 6000:
                return ShadingPatternValues.Percent60;
            case 6250:
                return ShadingPatternValues.Percent62;
            case 6500:
                return ShadingPatternValues.Percent65;
            case 7000:
                return ShadingPatternValues.Percent70;
            case 7500:
                return ShadingPatternValues.Percent75;
            case 8000:
                return ShadingPatternValues.Percent80;
            case 8500:
                return ShadingPatternValues.Percent85;
            case 8750:
                return ShadingPatternValues.Percent87;
            case 9000:
                return ShadingPatternValues.Percent90;
            case 9500:
                return ShadingPatternValues.Percent95;
            case 10000:
                return ShadingPatternValues.Solid;
            default:
                return null;
        }
    }

    private static ShadingPatternValues? ToWordShadingPattern(RtfShadingPattern pattern) {
        switch (pattern) {
            case RtfShadingPattern.Horizontal:
                return ShadingPatternValues.ThinHorizontalStripe;
            case RtfShadingPattern.Vertical:
                return ShadingPatternValues.ThinVerticalStripe;
            case RtfShadingPattern.ForwardDiagonal:
                return ShadingPatternValues.ThinDiagonalStripe;
            case RtfShadingPattern.BackwardDiagonal:
                return ShadingPatternValues.ThinReverseDiagonalStripe;
            case RtfShadingPattern.Cross:
                return ShadingPatternValues.ThinHorizontalCross;
            case RtfShadingPattern.DiagonalCross:
                return ShadingPatternValues.ThinDiagonalCross;
            case RtfShadingPattern.DarkHorizontal:
                return ShadingPatternValues.HorizontalStripe;
            case RtfShadingPattern.DarkVertical:
                return ShadingPatternValues.VerticalStripe;
            case RtfShadingPattern.DarkForwardDiagonal:
                return ShadingPatternValues.DiagonalStripe;
            case RtfShadingPattern.DarkBackwardDiagonal:
                return ShadingPatternValues.ReverseDiagonalStripe;
            case RtfShadingPattern.DarkCross:
                return ShadingPatternValues.HorizontalCross;
            case RtfShadingPattern.DarkDiagonalCross:
                return ShadingPatternValues.DiagonalCross;
            default:
                return null;
        }
    }

    private static void ApplyParagraphBorder(RtfParagraphBorder source, Action<BorderValues?> setStyle, Action<UInt32Value?> setWidth, Action<string?> setColor, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        setStyle(ToWordParagraphBorderStyle(source.Style));
        if (source.Width.HasValue && source.Width.Value >= 0) {
            setWidth((UInt32Value)(uint)source.Width.Value);
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                setColor(color);
            }
        }
    }

    private static RtfParagraphBorderStyle ToRtfParagraphBorderStyle(BorderValues? value) {
        if (value == BorderValues.Double) return RtfParagraphBorderStyle.Double;
        if (value == BorderValues.Dotted) return RtfParagraphBorderStyle.Dotted;
        if (value == BorderValues.Dashed) return RtfParagraphBorderStyle.Dashed;
        if (value == BorderValues.Nil || value == BorderValues.None) return RtfParagraphBorderStyle.None;
        if (value == BorderValues.Single) return RtfParagraphBorderStyle.Single;
        return RtfParagraphBorderStyle.None;
    }

    private static BorderValues? ToWordParagraphBorderStyle(RtfParagraphBorderStyle value) {
        switch (value) {
            case RtfParagraphBorderStyle.Double:
                return BorderValues.Double;
            case RtfParagraphBorderStyle.Dotted:
                return BorderValues.Dotted;
            case RtfParagraphBorderStyle.Dashed:
                return BorderValues.Dashed;
            case RtfParagraphBorderStyle.Single:
                return BorderValues.Single;
            default:
                return BorderValues.Nil;
        }
    }
}
