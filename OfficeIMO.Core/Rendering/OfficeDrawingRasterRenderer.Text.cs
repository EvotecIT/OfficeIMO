using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderTransformedPositionedText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale, long maximumRasterPixels) {
        canvas.CancellationToken.ThrowIfCancellationRequested();
        _ = OfficeRasterExportPlanner.Resolve(text.Width, text.Height, OfficeImageExportFormat.Png,
            new OfficeImageExportOptions { Scale = scale, MaximumRasterPixels = maximumRasterPixels, RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw });
        var layer = new OfficeRasterImage(Math.Max(1, (int)Math.Ceiling(text.Width * scale)), Math.Max(1, (int)Math.Ceiling(text.Height * scale)));
        var local = new OfficeRasterCanvas(layer, font: canvas.OutlineFont, fonts: canvas.Fonts,
            textShapingProvider: canvas.TextShapingProvider, textShapingLanguage: canvas.TextShapingLanguage,
            diagnosticSink: canvas.DiagnosticSink, diagnosticSource: canvas.DiagnosticSource, cancellationToken: canvas.CancellationToken);
        local.DrawPositionedText(text.Text, 0D, text.BaselineOffset * scale, text.Width * scale, text.Height * scale,
            text.Color ?? OfficeColor.Black, Math.Max(1D, text.Font.Size * scale) * text.BaselineScale,
            text.Alignment, text.Font.Style, text.Font.FamilyName, text.TextAdvanceWidth!.Value * scale,
            text.UnderlineStyle, text.StrikethroughStyle, text.DecorationColor, text.FeatureSettings, text.FontPalette,
            baselineFontSize: Math.Max(1D, text.Font.Size * scale));
        var frame = new OfficeImageFrameTransform(text.RotationDegrees, text.RotationCenterX * scale, text.RotationCenterY * scale,
            text.FlipHorizontal, text.FlipVertical);
        OfficeTransform transform = OfficeTransform.Translate(text.X * scale, text.Y * scale).Then(frame.CreateDestinationTransform());
        canvas.DrawAffineImage(layer, transform, 1D, OfficeBlendMode.Normal, interpolate: true);
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
            double positionedSourceFontSize = Math.Max(1D, text.Font.Size * scale);
            double renderedFontSize = positionedSourceFontSize * text.BaselineScale;
            double positionedBaselineOffset = text.BaselineOffset * scale;
            canvas.DrawPositionedText(
                text.Text,
                contentX,
                contentY + positionedBaselineOffset,
                contentWidth,
                contentHeight,
                text.Color ?? OfficeColor.Black,
                renderedFontSize,
                text.Alignment,
                text.Font.Style,
                text.Font.FamilyName,
                (text.TextAdvanceWidth ?? canvas.MeasureText(text.Text, renderedFontSize, text.Font.FamilyName, text.Font.Style) / scale) * scale,
                text.UnderlineStyle,
                text.StrikethroughStyle,
                text.DecorationColor,
                text.FeatureSettings,
                text.FontPalette,
                baselineFontSize: positionedSourceFontSize);
            return;
        }

        if (supportsLegacyFastPath && supportsPositionedPath) {
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
        Func<string?, double, double> measure = (value, size) => canvas.MeasureText(value, size, text.Font.FamilyName);
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
                paragraphIndent)
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
