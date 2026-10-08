using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private const long MaximumSingleTransformedTextIntermediatePixels = 16_000_000L;

    private static void RenderTransformedPositionedText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale, long maximumRasterPixels) {
        canvas.CancellationToken.ThrowIfCancellationRequested();
        using var cffScope = canvas.PushCffExecutionScope();
        OfficeImageFrameTransform frame = new OfficeImageFrameTransform(text.RotationDegrees, text.RotationCenterX * scale,
            text.RotationCenterY * scale, text.FlipHorizontal, text.FlipVertical);
        (double axisX, double axisY) = GetTextLayerAxisScales(frame, canvas);
        double left = 0D, top = 0D, right = text.Width * scale, bottom = text.Height * scale;
        double sourceSize = Math.Max(.1D, text.Font.Size * scale);
        string[] lines = text.RasterText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scale;
        for (int index = 0; index < lines.Length; index++) {
            double offset = index * lineHeight;
            if (offset >= text.Height * scale) break;
            if (lines[index].Length == 0) continue;
            double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue
                ? text.TextAdvanceWidth.Value * scale
                : Math.Max(.001D, canvas.MeasurePositionedText(lines[index], sourceSize * text.BaselineScale,
                    text.Font.FamilyName, text.Font.Style, text.FeatureSettings, text.TextDirection));
            var bounds = canvas.MeasurePositionedTextBounds(lines[index], 0D, offset + text.BaselineOffset * scale,
                text.Width * scale, text.Height * scale - offset, sourceSize * text.BaselineScale, text.Font, advance,
                text.Alignment, text.FeatureSettings, text.FontPalette, sourceSize, text.UnderlineStyle, text.StrikethroughStyle,
                text.TextDirection);
            left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
            right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
        }
        left = Math.Floor(left); top = Math.Floor(top);
        right = Math.Ceiling(right); bottom = Math.Ceiling(bottom);
        // Transparent sampling support is optional; actual ink and the caller's pixel
        // ceiling are not. An exact-fit frame must not fail solely because of padding.
        double paddedWidth = Math.Max(1D, Math.Ceiling((right - left + 2D) * axisX));
        double paddedHeight = Math.Max(1D, Math.Ceiling((bottom - top + 2D) * axisY));
        long remainingPixels = Math.Min(MaximumSingleTransformedTextIntermediatePixels,
            canvas.GetRemainingTransformedTextIntermediatePixels(maximumRasterPixels));
        if (paddedHeight > 0D && paddedWidth <= remainingPixels / paddedHeight) {
            left -= 1D; top -= 1D; right += 1D; bottom += 1D;
        }
        double pixelWidth = Math.Max(1D, Math.Ceiling((right - left) * axisX));
        double pixelHeight = Math.Max(1D, Math.Ceiling((bottom - top) * axisY));
        OfficeTransform transform = OfficeTransform.Scale(1D / axisX, 1D / axisY)
            .Then(OfficeTransform.Translate(text.X * scale + left, text.Y * scale + top)).Then(frame.CreateDestinationTransform());
        // Use the measured ink, padding and rounded layer, not just the text frame.
        if (!double.IsInfinity(pixelWidth) && !double.IsInfinity(pixelHeight) &&
            !canvas.IntersectsVisibleSurface(transform, pixelWidth, pixelHeight)) return;
        _ = OfficeRasterExportPlanner.Resolve(pixelWidth, pixelHeight, OfficeImageExportFormat.Png,
            new OfficeImageExportOptions { MaximumRasterPixels = Math.Min(maximumRasterPixels, MaximumSingleTransformedTextIntermediatePixels), RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw });
        canvas.ChargeTransformedTextIntermediatePixels(
            checked((long)pixelWidth * (long)pixelHeight), maximumRasterPixels);
        OfficeRasterImage layer = new OfficeRasterImage((int)pixelWidth, (int)pixelHeight);
        var local = new OfficeRasterCanvas(layer, font: canvas.OutlineFont, fonts: canvas.Fonts,
            textShapingProvider: canvas.TextShapingProvider, textShapingLanguage: canvas.TextShapingLanguage,
            diagnosticSink: canvas.DiagnosticSink, diagnosticSource: canvas.DiagnosticSource, cancellationToken: canvas.CancellationToken);
        local.FontMetricScale = canvas.FontMetricScale;
        local.SetCoordinateScale(axisX, axisY);
        using var faceScope = local.PushTextFace(text.Font.Face);
        local.ShareCffOperationBudget(canvas);
        local.ShareTransformedTextBudget(canvas.TransformedTextBudget);
        local.PreservePaintedGlyphOrder = canvas.PreservePaintedGlyphOrder;
        RenderPositionedTextLines(local, text, scale, -left, -top, text.Width * scale, text.Height * scale);
        canvas.DrawAffineImage(layer, transform, 1D, OfficeBlendMode.Normal, interpolate: true);
    }

    private static void RenderPositionedTextLines(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale,
        double x, double y, double width, double height) {
        double sourceSize = Math.Max(.1D, text.Font.Size * scale);
        double size = sourceSize * text.BaselineScale;
        double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scale;
        string[] lines = text.RasterText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            double offset = index * lineHeight;
            if (offset >= height) break;
            string value = lines[index];
            if (value.Length == 0) continue;
            double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue
                ? text.TextAdvanceWidth.Value * scale
                : Math.Max(.001D, canvas.MeasurePositionedText(value, size, text.Font.FamilyName,
                    text.Font.Style, text.FeatureSettings, text.TextDirection));
            canvas.DrawPositionedText(value, x, y + offset + text.BaselineOffset * scale, width, height - offset,
                text.Color ?? OfficeColor.Black, size, text.Alignment, text.Font.Style, text.Font.FamilyName, advance,
                text.UnderlineStyle, text.StrikethroughStyle, text.DecorationColor, text.FeatureSettings, text.FontPalette,
                baselineFontSize: sourceSize, textDirection: text.TextDirection);
        }
    }

    private static bool TryRenderTransformedVerticalText(
        OfficeRasterCanvas canvas,
        OfficeDrawingText text,
        double scale,
        double contentX,
        double contentY,
        double contentWidth,
        double contentHeight,
        long maximumRasterPixels) {
        OfficeImageFrameTransform frame = new OfficeImageFrameTransform(text.RotationDegrees,
            text.RotationCenterX * scale, text.RotationCenterY * scale, text.FlipHorizontal, text.FlipVertical);
        (double axisX, double axisY) = GetTextLayerAxisScales(frame, canvas);
        double pixelWidth = Math.Max(1D, Math.Ceiling(contentWidth * axisX));
        double pixelHeight = Math.Max(1D, Math.Ceiling(contentHeight * axisY));
        OfficeTransform transform = OfficeTransform.Scale(1D / axisX, 1D / axisY)
            .Then(CreateVerticalTextPlacement(text, scale, contentX, contentY));
        if (!double.IsInfinity(pixelWidth) && !double.IsInfinity(pixelHeight) &&
            !canvas.IntersectsVisibleSurface(transform, pixelWidth, pixelHeight)) {
            // Unsupported vertical positioning still reaches the stacked fallback,
            // whose ink may overhang the frame that was just culled.
            return (text.Color ?? OfficeColor.Black).A == 0 || canvas.CanDrawVerticalText(text.RasterText,
                text.Font.Size * scale, text.Font.Style, text.Font.FamilyName, text.FeatureSettings);
        }
        _ = OfficeRasterExportPlanner.Resolve(pixelWidth, pixelHeight, OfficeImageExportFormat.Png,
            new OfficeImageExportOptions {
                MaximumRasterPixels = Math.Min(maximumRasterPixels, MaximumSingleTransformedTextIntermediatePixels),
                RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
            });
        long layerPixels = checked((long)pixelWidth * (long)pixelHeight);
        canvas.ChargeTransformedTextIntermediatePixels(layerPixels, maximumRasterPixels);
        OfficeRasterImage layer = new OfficeRasterImage((int)pixelWidth, (int)pixelHeight);
        var local = new OfficeRasterCanvas(layer, font: canvas.OutlineFont, fonts: canvas.Fonts,
            textShapingProvider: canvas.TextShapingProvider, textShapingLanguage: canvas.TextShapingLanguage,
            diagnosticSink: canvas.DiagnosticSink, diagnosticSource: canvas.DiagnosticSource,
            cancellationToken: canvas.CancellationToken);
        local.FontMetricScale = canvas.FontMetricScale;
        local.SetCoordinateScale(axisX, axisY);
        using var faceScope = local.PushTextFace(text.Font.Face);
        local.ShareCffOperationBudget(canvas);
        local.ShareTransformedTextBudget(canvas.TransformedTextBudget);
        local.PreservePaintedGlyphOrder = canvas.PreservePaintedGlyphOrder;
        using (local.PushClipRectangle(0D, 0D, contentWidth, contentHeight)) {
            if (!local.TryDrawVerticalText(
                text.RasterText,
                0D,
                0D,
                contentWidth,
                contentHeight,
                text.Color ?? OfficeColor.Black,
                text.Font.Size * scale,
                text.Font.Style,
                text.Font.FamilyName,
                text.FeatureSettings,
                text.FontPalette,
                text.UnderlineStyle,
                text.StrikethroughStyle,
                text.DecorationColor)) {
                canvas.ReleaseTransformedTextIntermediatePixels(layerPixels);
                return false;
            }
        }

        canvas.DrawAffineImage(layer, transform, 1D, OfficeBlendMode.Normal, interpolate: true);
        return true;
    }

    private static (double X, double Y) GetTextLayerAxisScales(OfficeImageFrameTransform frame, OfficeRasterCanvas canvas) {
        // Text frames contain only rotation/reflection. Preserve their exact unit
        // density on isotropic canvases rather than letting roundoff add a pixel.
        return canvas.CoordinateScaleX == canvas.CoordinateScaleY
            ? (canvas.CoordinateScaleX, canvas.CoordinateScaleY)
            : GetEffectAxisScales(frame.CreateDestinationTransform(), canvas.CoordinateScaleX, canvas.CoordinateScaleY);
    }

    internal static (double X, double Y, double Width, double Height) ResolveTextContentRectangle(OfficeDrawingText text, double scale) {
        OfficeTextPadding padding = text.Padding.Scale(scale);
        return ((text.X * scale) + padding.Left, (text.Y * scale) + padding.Top,
            (text.Width * scale) - padding.Horizontal, (text.Height * scale) - padding.Vertical);
    }

    internal static OfficeTransform CreateVerticalTextPlacement(OfficeDrawingText text, double scale, double contentX, double contentY) {
        var frame = new OfficeImageFrameTransform(text.RotationDegrees, text.RotationCenterX * scale,
            text.RotationCenterY * scale, text.FlipHorizontal, text.FlipVertical);
        return OfficeTransform.Translate(contentX, contentY).Then(frame.CreateDestinationTransform());
    }

    internal static void RenderText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale, long maximumRasterPixels) {
        using var metricScope = canvas.PushFontMetricScale(canvas.FontMetricScale * text.FontMetricScale);
        using var faceScope = canvas.PushTextFace(text.Font.Face);
        var (contentX, contentY, contentWidth, contentHeight) = ResolveTextContentRectangle(text, scale);
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        if (text.TextDirection == OfficeTextDirection.TopToBottom) {
            if (text.HasFrameTransform) {
                if (TryRenderTransformedVerticalText(
                    canvas,
                    text,
                    scale,
                    contentX,
                    contentY,
                    contentWidth,
                    contentHeight,
                    maximumRasterPixels)) {
                    return;
                }
            } else {
                using (canvas.PushClipRectangle(contentX, contentY, contentWidth, contentHeight)) {
                    if (canvas.TryDrawVerticalText(
                        text.RasterText,
                        contentX,
                        contentY,
                        contentWidth,
                        contentHeight,
                        text.Color ?? OfficeColor.Black,
                        text.Font.Size * scale,
                        text.Font.Style,
                        text.Font.FamilyName,
                        text.FeatureSettings,
                        text.FontPalette,
                        text.UnderlineStyle,
                        text.StrikethroughStyle,
                        text.DecorationColor)) {
                        return;
                    }
                }
            }
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
             text.OverflowBehavior == OfficeTextOverflowBehavior.Clip ||
             text.BaselineScale != 1D || text.BaselineOffset != 0D ||
             !text.FeatureSettings.IsDefault ||
             !string.Equals(text.FontPalette, "normal", StringComparison.OrdinalIgnoreCase)) && supportsPositionedPath) {
            RenderPositionedTextLines(canvas, text, scale, contentX, contentY, contentWidth, contentHeight);
            return;
        }

        if (supportsLegacyFastPath && supportsPositionedPath && text.RasterText.IndexOfAny(new[] { '\r', '\n' }) < 0) {
            canvas.DrawBaselineText(
                text.RasterText,
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

        double sourceFontSize = Math.Max(.1D, text.Font.Size * scale);
        double fontSize = sourceFontSize * text.BaselineScale;
        double baselineOffset = text.BaselineOffset * scale;
        OfficeTextParagraphIndent paragraphIndent = text.ParagraphIndent.Scale(scale);
        double lineHeightFactor = OfficeDrawingTextLayout.ResolveLineHeightFactor(text.LineHeight * scale, fontSize);
        double minimumFontSize = Math.Min(6D * scale, fontSize);
        Func<string?, double, double> measure = (value, size) => canvas.MeasureText(value, size, text.Font.FamilyName, text.Font.Style);
        OfficeTextBlockLayout layout = text.StackedText
            ? OfficeTextLayoutEngine.LayoutStackedTextBlockCore(
                text.RasterText,
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
                text.RasterText,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                paragraphIndent,
                (value, size) => canvas.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.RasterText,
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
