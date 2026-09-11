using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendText(StringBuilder sb, OfficeDrawingText text, OfficeRasterCanvas measurement) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical;
        if (useFrameTransform) {
            AppendTextFrameGroupStart(sb, text);
        }

        if (text.WrapText || text.ShrinkToFit || text.StackedText || text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.HasPadding) {
            AppendTextBlock(sb, text, measurement, useFrameTransform);
            if (useFrameTransform) {
                sb.Append("</g>");
            }

            return;
        }

        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double x = contentX;
        if (text.Alignment == OfficeTextAlignment.Center) {
            x += contentWidth / 2D;
        } else if (text.Alignment == OfficeTextAlignment.Right) {
            x += contentWidth;
        }

        double sourceFontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double fontSize = sourceFontSize * text.BaselineScale;
        double y = contentY + sourceFontSize + text.BaselineOffset;
        double lineHeight = text.LineHeight ?? sourceFontSize * 1.2D;
        if (text.TextAdvanceWidth.HasValue) {
            sb.AppendSvgPositionedTextElement(
                text.Text,
                x,
                y,
                lineHeight,
                text.Color ?? OfficeColor.Black,
                text.Font.FamilyName ?? "Arial",
                fontSize,
                text.Alignment,
                text.Font.IsBold,
                text.Font.IsItalic,
                (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
                useFrameTransform ? 0D : text.RotationDegrees,
                useFrameTransform ? 0D : text.RotationCenterX,
                useFrameTransform ? 0D : text.RotationCenterY,
                (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough,
                text.TextAdvanceWidth.Value,
                text.UnderlineStyle,
                text.StrikethroughStyle,
                OfficeTextBaseline.Normal,
                text.DecorationColor,
                text.FeatureSettings,
                text.FontPalette);
        } else {
            sb.AppendSvgFeaturedTextElement(
                text.Text,
                x,
                y,
                lineHeight,
                text.Color ?? OfficeColor.Black,
                text.Font.FamilyName ?? "Arial",
                fontSize,
                text.Alignment,
                text.Font.IsBold,
                text.Font.IsItalic,
                (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
                useFrameTransform ? 0D : text.RotationDegrees,
                useFrameTransform ? 0D : text.RotationCenterX,
                useFrameTransform ? 0D : text.RotationCenterY,
                (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough,
                text.UnderlineStyle,
                text.StrikethroughStyle,
                OfficeTextBaseline.Normal,
                text.DecorationColor,
                text.FeatureSettings,
                text.FontPalette);
        }

        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static void AppendTextBlock(StringBuilder sb, OfficeDrawingText text, OfficeRasterCanvas measurement, bool useFrameTransform = false) {
        double sourceFontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double fontSize = sourceFontSize * text.BaselineScale;
        double baselineOffset = text.BaselineOffset;
        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, text.LineHeight.Value / fontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D, fontSize);
        Func<string?, double, double> measure = (value, size) => measurement.MeasureText(value, size, text.Font.FamilyName, text.Font.Style);
        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        OfficeTextBlockLayout layout = text.StackedText
            ? OfficeTextLayoutEngine.LayoutStackedTextBlock(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ShrinkToFit)
            : text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedText(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ParagraphIndent)
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                wrap: text.WrapText,
                shrinkToFit: text.ShrinkToFit,
                paragraphIndent: text.ParagraphIndent);
        sb.AppendSvgStyledTextBlock(
            layout,
            contentX,
            contentY + baselineOffset,
            contentWidth,
            contentHeight,
            text.Color ?? OfficeColor.Black,
            string.IsNullOrWhiteSpace(text.Font.FamilyName) ? "Arial" : text.Font.FamilyName,
            text.Alignment,
            text.VerticalAlignment,
            text.Font.IsBold,
            text.Font.IsItalic,
            (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
            useFrameTransform ? 0D : text.RotationDegrees,
            useFrameTransform ? 0D : text.RotationCenterX,
            useFrameTransform ? 0D : text.RotationCenterY,
            centerLineInLineHeight: true,
            strikethrough: (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough,
            underlineStyle: text.UnderlineStyle,
            strikethroughStyle: text.StrikethroughStyle,
            baseline: OfficeTextBaseline.Normal,
            decorationColor: text.DecorationColor);
    }

    private static void AppendRichText(StringBuilder sb, OfficeDrawingRichText text, OfficeRasterCanvas measurement) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical;
        if (useFrameTransform) {
            AppendRichTextFrameGroupStart(sb, text);
        }

        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            if (useFrameTransform) {
                sb.Append("</g>");
            }

            return;
        }

        OfficeRichTextBlockLayout layout = CreateRichTextLayout(text, contentWidth, contentHeight, measurement);
        sb.AppendSvgRichTextBlock(
            layout,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            text.Alignment,
            text.VerticalAlignment,
            useFrameTransform ? 0D : text.RotationDegrees,
            useFrameTransform ? 0D : text.RotationCenterX,
            useFrameTransform ? 0D : text.RotationCenterY);
        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static OfficeRichTextBlockLayout CreateRichTextLayout(OfficeDrawingRichText text, double contentWidth, double contentHeight, OfficeRasterCanvas measurement) {
        double maxFontSize = 10D;
        for (int i = 0; i < text.Runs.Count; i++) {
            maxFontSize = Math.Max(maxFontSize, text.Runs[i].FontSize);
        }

        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, text.LineHeight.Value / maxFontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D, maxFontSize);
        Func<string?, double, string?, double> measure = (value, size, family) => measurement.MeasureText(value, size, family);
        return OfficeTextLayoutEngine.LayoutRichTextBlock(
            text.Runs,
            contentWidth,
            contentHeight,
            lineHeightFactor,
            measure,
            text.WrapText,
            text.ShrinkToFit,
            minimumFontSize,
            text.ParagraphIndent);
    }

    private static void AppendTextFrameGroupStart(StringBuilder sb, OfficeDrawingText text) {
        string? transform = OfficeSvgFormatting.FormatImageFrameTransform(text.CreateFrameTransform());
        if (string.IsNullOrWhiteSpace(transform)) {
            sb.Append("<g>");
            return;
        }

        sb.Append("<g")
            .AppendAttribute("transform", transform)
            .Append('>');
    }

}
