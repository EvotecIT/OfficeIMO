using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendText(
        StringBuilder sb,
        OfficeDrawingText text,
        OfficeRasterCanvas textMetrics,
        string idPrefix,
        ref int clipPathId,
        bool textAsPaths, System.Collections.Generic.ISet<string> requiredFontFamilies) {
        var fragment = new StringBuilder(Math.Min(256, sb.MaxCapacity - sb.Length), sb.MaxCapacity - sb.Length);
        AppendTextCore(fragment, text, textMetrics, idPrefix, ref clipPathId, textAsPaths);
        CollectSvgTextFontFamilies(fragment.ToString(), requiredFontFamilies, textMetrics.CancellationToken);
        sb.Append(fragment);
    }

    private static void AppendTextCore(StringBuilder sb, OfficeDrawingText text, OfficeRasterCanvas textMetrics,
        string idPrefix, ref int clipPathId, bool textAsPaths) {
        using var metricScope = textMetrics.PushFontMetricScale(text.FontMetricScale);
        using var faceScope = textMetrics.PushTextFace(text.Font.Face);
        if (TryAppendGlyphPaths(sb, text, textMetrics, textAsPaths || text.GlyphTransform.HasValue)) return;
        if (text.GlyphTransform is { } glyphTransform) {
            if (text.OutlineStrokeColor.HasValue) throw new NotSupportedException(
                "Stroked affine glyphs require a resolved single-font outline run for SVG export.");
            // Filled text can retain the existing SVG font/fallback path when a
            // color or unsupported outline run cannot be emitted as contours.
            // A stroked run cannot use this post-paint transform: it would also
            // change its destination pen, so that unsupported profile is explicit.
            OfficeDrawingText local = text.Clone();
            local.GlyphTransform = null;
            if (text.HasFrameTransform) {
                OfficeTransform frame = text.CreateFrameTransform().CreateDestinationTransform();
                glyphTransform = frame.Invert().Then(glyphTransform).Then(frame);
            }
            sb.Append("<g").Append(BuildMatrixTransformAttribute(glyphTransform, 0D, 0D)).Append('>');
            AppendTextCore(sb, local, textMetrics, idPrefix, ref clipPathId, textAsPaths);
            sb.Append("</g>");
            return;
        }
        if (text.OutlineStrokeColor.HasValue) {
            var fragment = new StringBuilder(Math.Min(256, sb.MaxCapacity - sb.Length), sb.MaxCapacity - sb.Length);
            AppendTextContent(fragment, text, textMetrics, idPrefix, ref clipPathId);
            var painted = new StringBuilder();
            if (textMetrics.Fonts is { } strokeFonts && HasSelectedVariableFont(text, strokeFonts)) {
                AppendSelectedVariableFonts(painted, fragment.ToString(), text, strokeFonts, textMetrics.CancellationToken);
            } else painted.Append(fragment);
            AppendStrokedSvgText(sb, painted.ToString(), text, textMetrics.CancellationToken);
        } else if (textMetrics.Fonts is { } fonts && HasSelectedVariableFont(text, fonts)) {
            var fragment = new StringBuilder();
            AppendTextContent(fragment, text, textMetrics, idPrefix, ref clipPathId);
            AppendSelectedVariableFonts(sb, fragment.ToString(), text, fonts, textMetrics.CancellationToken);
        } else AppendTextContent(sb, text, textMetrics, idPrefix, ref clipPathId);
    }

    private static void AppendTextContent(StringBuilder sb, OfficeDrawingText text,
        OfficeRasterCanvas textMetrics, string idPrefix, ref int clipPathId) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical ||
            (text.TextDirection == OfficeTextDirection.TopToBottom && Math.Abs(text.RotationDegrees) > 0.000001D);
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
                text.RasterText,
                verticalContentX + verticalContentWidth / 2D,
                verticalContentY,
                ResolveSvgTextFill(text),
                text.Font.FamilyName,
                text.Font.Size,
                text.Font.IsBold,
                text.Font.IsItalic,
                text.Font.IsUnderline,
                text.Font.IsStrikethrough,
                text.UnderlineStyle,
                text.StrikethroughStyle,
                text.DecorationColor,
                text.FeatureSettings,
                text.FontPalette, text.Font.Face);
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
            text.RasterText,
            x,
            y,
            lineHeight,
            ResolveSvgTextFill(text),
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
            OfficeTextShapingBackend.BrowserNative,
            text.PreservesPaintedGlyphs ? OfficeTextDirection.LeftToRight : text.TextDirection,
            text.PreservesPaintedGlyphs, face: text.Font.Face, logicalText: text.Text);

        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static void AppendTextBlock(StringBuilder sb, OfficeDrawingText text, OfficeRasterCanvas textMetrics, bool useFrameTransform = false) {
        double sourceFontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double fontSize = sourceFontSize * text.BaselineScale;
        double baselineOffset = text.BaselineOffset;
        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        OfficeTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, contentWidth, contentHeight, textMetrics);
        sb.AppendSvgStyledTextBlockWithFace(
            layout,
            text.Font.Face,
            contentX,
            contentY + baselineOffset,
            contentWidth,
            contentHeight,
            ResolveSvgTextFill(text),
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

    private static void AppendRichText(StringBuilder sb, OfficeDrawingRichText text, OfficeRasterCanvas textMetrics,
        System.Collections.Generic.ISet<string> requiredFontFamilies) {
        var fragment = new StringBuilder(Math.Min(256, sb.MaxCapacity - sb.Length), sb.MaxCapacity - sb.Length);
        AppendRichTextCore(fragment, text, textMetrics);
        CollectSvgTextFontFamilies(fragment.ToString(), requiredFontFamilies, textMetrics.CancellationToken);
        sb.Append(fragment);
    }

    private static void AppendRichTextCore(StringBuilder sb, OfficeDrawingRichText text, OfficeRasterCanvas textMetrics) {
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

        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, contentWidth, contentHeight, textMetrics);
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
