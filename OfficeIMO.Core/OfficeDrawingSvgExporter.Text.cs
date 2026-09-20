using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendText(
        StringBuilder sb,
        OfficeDrawingText text,
        OfficeRasterCanvas textMetrics,
        string idPrefix,
        ref int clipPathId) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical;
        if (useFrameTransform) {
            AppendTextFrameGroupStart(sb, text);
        }

        if (text.TextDirection == OfficeTextDirection.TopToBottom) {
            double verticalContentX = text.X + text.Padding.Left;
            double verticalContentY = text.Y + text.Padding.Top;
            double verticalContentWidth = text.Width - text.Padding.Horizontal;
            double verticalContentHeight = text.Height - text.Padding.Vertical;
            if (verticalContentWidth <= 0D || verticalContentHeight <= 0D) {
                if (useFrameTransform) sb.Append("</g>");
                return;
            }

            string verticalClipPathId = idPrefix + "officeimo-text-clip-" +
                (++clipPathId).ToString(System.Globalization.CultureInfo.InvariantCulture);
            sb.Append("<defs><clipPath id=\"")
                .Append(verticalClipPathId)
                .Append("\"><rect x=\"").Append(Format(verticalContentX))
                .Append("\" y=\"").Append(Format(verticalContentY))
                .Append("\" width=\"").Append(Format(verticalContentWidth))
                .Append("\" height=\"").Append(Format(verticalContentHeight))
                .Append("\"/></clipPath></defs><g")
                .AppendClipPathReference(verticalClipPathId)
                .Append('>');
            sb.AppendSvgVerticalTextElement(
                text.Text,
                verticalContentX + verticalContentWidth / 2D,
                verticalContentY,
                text.Color ?? OfficeColor.Black,
                text.Font.FamilyName,
                text.Font.Size,
                text.Font.IsBold,
                text.Font.IsItalic,
                text.FeatureSettings,
                text.FontPalette);
            sb.Append("</g>");
            if (useFrameTransform) sb.Append("</g>");
            return;
        }

        if (text.WrapText || text.ShrinkToFit || text.StackedText || text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.HasPadding) {
            AppendTextBlock(sb, text, textMetrics, useFrameTransform);
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
        double? advance = text.TextAdvanceWidth;
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
            advance,
            text.UnderlineStyle,
            text.StrikethroughStyle,
            OfficeTextBaseline.Normal,
            text.DecorationColor,
            text.FeatureSettings,
            text.FontPalette,
            OfficeTextShapingBackend.BrowserNative);

        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static void AppendTextBlock(StringBuilder sb, OfficeDrawingText text, OfficeRasterCanvas textMetrics, bool useFrameTransform = false) {
        double sourceFontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double fontSize = sourceFontSize * text.BaselineScale;
        double baselineOffset = text.BaselineOffset;
        double lineHeightFactor = OfficeDrawingTextLayout.ResolveLineHeightFactor(text.LineHeight, fontSize);
        double minimumFontSize = Math.Min(6D, fontSize);
        Func<string?, double, double> measure = (value, size) =>
            textMetrics.MeasureText(value, size, text.Font.FamilyName, text.Font.Style);
        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        OfficeTextBlockLayout layout = text.StackedText
            ? OfficeTextLayoutEngine.LayoutStackedTextBlockCore(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ShrinkToFit,
                (value, size) => textMetrics.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedTextCore(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ParagraphIndent,
                (value, size) => textMetrics.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
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
            text.Font.FamilyName,
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

    private static void AppendRichText(StringBuilder sb, OfficeDrawingRichText text, OfficeRasterCanvas textMetrics) {
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

        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.Create(text, contentWidth, contentHeight, textMetrics.MeasureText, measurePaint: textMetrics.MeasureTextPaintBounds);
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
