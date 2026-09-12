using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderTransformedPositionedText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale, long maximumRasterPixels) {
        canvas.CancellationToken.ThrowIfCancellationRequested();
        double left = 0D, top = 0D, right = text.Width * scale, bottom = text.Height * scale;
        double sourceSize = Math.Max(1D, text.Font.Size * scale);
        string[] lines = text.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scale;
        for (int index = 0; index < lines.Length; index++) {
            double offset = index * lineHeight;
            if (offset >= text.Height * scale) break;
            if (lines[index].Length == 0) continue;
            double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue
                ? text.TextAdvanceWidth.Value * scale
                : Math.Max(.001D, canvas.MeasureText(lines[index], sourceSize * text.BaselineScale, text.Font.FamilyName, text.Font.Style));
            var bounds = canvas.MeasurePositionedTextBounds(lines[index], 0D, offset + text.BaselineOffset * scale,
                text.Width * scale, text.Height * scale - offset, sourceSize * text.BaselineScale, text.Font, advance,
                text.Alignment, text.FeatureSettings, text.FontPalette, sourceSize, text.UnderlineStyle, text.StrikethroughStyle);
            left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
            right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
        }
        left = Math.Floor(left); top = Math.Floor(top);
        right = Math.Ceiling(right); bottom = Math.Ceiling(bottom);
        // Transparent sampling support is optional; actual ink and the caller's pixel
        // ceiling are not. An exact-fit frame must not fail solely because of padding.
        double paddedWidth = right - left + 2D, paddedHeight = bottom - top + 2D;
        if (paddedHeight > 0D && paddedWidth <= maximumRasterPixels / paddedHeight) {
            left -= 1D; top -= 1D; right += 1D; bottom += 1D;
        }
        _ = OfficeRasterExportPlanner.Resolve(right - left, bottom - top, OfficeImageExportFormat.Png,
            new OfficeImageExportOptions { MaximumRasterPixels = maximumRasterPixels, RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw });
        var layer = new OfficeRasterImage(Math.Max(1, (int)(right - left)), Math.Max(1, (int)(bottom - top)));
        var local = new OfficeRasterCanvas(layer, font: canvas.OutlineFont, fonts: canvas.Fonts,
            textShapingProvider: canvas.TextShapingProvider, textShapingLanguage: canvas.TextShapingLanguage,
            diagnosticSink: canvas.DiagnosticSink, diagnosticSource: canvas.DiagnosticSource, cancellationToken: canvas.CancellationToken);
        RenderPositionedTextLines(local, text, scale, -left, -top, text.Width * scale, text.Height * scale);
        var frame = new OfficeImageFrameTransform(text.RotationDegrees, text.RotationCenterX * scale, text.RotationCenterY * scale,
            text.FlipHorizontal, text.FlipVertical);
        OfficeTransform transform = OfficeTransform.Translate(text.X * scale + left, text.Y * scale + top).Then(frame.CreateDestinationTransform());
        canvas.DrawAffineImage(layer, transform, 1D, OfficeBlendMode.Normal, interpolate: true);
    }

    private static void RenderPositionedTextLines(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale,
        double x, double y, double width, double height) {
        double sourceSize = Math.Max(1D, text.Font.Size * scale);
        double size = sourceSize * text.BaselineScale;
        double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scale;
        string[] lines = text.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            double offset = index * lineHeight;
            if (offset >= height) break;
            string value = lines[index];
            if (value.Length == 0) continue;
            double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue
                ? text.TextAdvanceWidth.Value * scale
                : Math.Max(.001D, canvas.MeasureText(value, size, text.Font.FamilyName, text.Font.Style));
            canvas.DrawPositionedText(value, x, y + offset + text.BaselineOffset * scale, width, height - offset,
                text.Color ?? OfficeColor.Black, size, text.Alignment, text.Font.Style, text.Font.FamilyName, advance,
                text.UnderlineStyle, text.StrikethroughStyle, text.DecorationColor, text.FeatureSettings, text.FontPalette,
                baselineFontSize: sourceSize);
        }
    }

    private static void RenderText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale, long maximumRasterPixels) {
        OfficeTextPadding scaledPadding = text.Padding.Scale(scale);
        double contentX = (text.X * scale) + scaledPadding.Left;
        double contentY = (text.Y * scale) + scaledPadding.Top;
        double contentWidth = (text.Width * scale) - scaledPadding.Horizontal;
        double contentHeight = (text.Height * scale) - scaledPadding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        if (text.HasFrameTransform && text.TextAdvanceWidth.HasValue && !text.WrapText && !text.ShrinkToFit &&
            !text.StackedText && !text.HasPadding && text.VerticalAlignment == OfficeTextVerticalAlignment.Top) {
            RenderTransformedPositionedText(canvas, text, scale, maximumRasterPixels);
            return;
        }

        bool supportsLegacyFastPath = Math.Abs(text.BaselineScale - 1D) < 0.000001D && Math.Abs(text.BaselineOffset) < 0.000001D &&
            text.UnderlineStyle == OfficeTextDecorationStyle.None &&
            text.StrikethroughStyle == OfficeTextDecorationStyle.None;
        bool supportsPositionedPath = !text.WrapText && !text.ShrinkToFit && !text.StackedText && !text.HasFrameTransform && text.VerticalAlignment == OfficeTextVerticalAlignment.Top && !text.HasPadding;
        if ((text.TextAdvanceWidth.HasValue ||
             text.BaselineScale != 1D || text.BaselineOffset != 0D ||
             !text.FeatureSettings.IsDefault ||
             !string.Equals(text.FontPalette, "normal", StringComparison.OrdinalIgnoreCase)) && supportsPositionedPath) {
            RenderPositionedTextLines(canvas, text, scale, contentX, contentY, contentWidth, contentHeight);
            return;
        }

        if (supportsLegacyFastPath && supportsPositionedPath && text.Text.IndexOfAny(new[] { '\r', '\n' }) < 0) {
            canvas.DrawText(
                text.Text,
                contentX,
                contentY,
                contentWidth,
                contentHeight,
                text.Color ?? OfficeColor.Black,
                text.Font.Size * scale,
                text.Alignment,
                text.Font.Style,
                text.Font.FamilyName);
            return;
        }

        double sourceFontSize = Math.Max(1D, text.Font.Size * scale);
        double fontSize = sourceFontSize * text.BaselineScale;
        double baselineOffset = text.BaselineOffset * scale;
        OfficeTextParagraphIndent paragraphIndent = text.ParagraphIndent.Scale(scale);
        double lineHeightFactor = OfficeDrawingTextLayout.ResolveLineHeightFactor(text.LineHeight * scale, fontSize);
        double minimumFontSize = Math.Min(6D * scale, fontSize);
        Func<string?, double, double> measure = (value, size) => canvas.MeasureText(value, size, text.Font.FamilyName, text.Font.Style);
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
                (value, size) => canvas.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedTextCore(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                paragraphIndent,
                (value, size) => canvas.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                wrap: text.WrapText,
                forceSingleLine: false,
                shrinkToFit: text.ShrinkToFit,
                overflowBehavior: text.OverflowBehavior,
                paragraphIndent: paragraphIndent);
        OfficeTextBlockRenderer.DrawRasterTextBlock(
            canvas,
            layout,
            contentX,
            contentY + baselineOffset,
            contentWidth,
            contentHeight,
            text.Color ?? OfficeColor.Black,
            text.Alignment,
            text.VerticalAlignment,
            (text.Font.Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold,
            (text.Font.Style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic,
            (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
            text.RotationDegrees,
            text.RotationCenterX * scale,
            text.RotationCenterY * scale,
            strikethrough: (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough,
            fontFamily: text.Font.FamilyName,
            flipHorizontal: text.FlipHorizontal,
            flipVertical: text.FlipVertical,
            underlineStyle: text.UnderlineStyle,
            strikethroughStyle: text.StrikethroughStyle,
            baseline: OfficeTextBaseline.Normal,
            decorationColor: text.DecorationColor);
    }

}
