using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free raster renderer for <see cref="OfficeDrawing"/> scenes.
/// </summary>
public static partial class OfficeDrawingRasterRenderer {
    /// <summary>
    /// Renders a drawing to an RGBA raster image.
    /// </summary>
    public static OfficeRasterImage Render(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) {
        return Render(drawing, new OfficeDrawingRasterRenderOptions { Scale = scale, Background = background });
    }

    /// <summary>Renders a drawing with an optional external image codec.</summary>
    public static OfficeRasterImage Render(OfficeDrawing drawing, OfficeDrawingRasterRenderOptions options) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (options == null) throw new ArgumentNullException(nameof(options));
        options.CancellationToken.ThrowIfCancellationRequested();
        double scale = options.Scale;

        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Scale must be a finite positive number.");
        }

        int width = Math.Max(1, (int)Math.Ceiling(drawing.Width * scale));
        int height = Math.Max(1, (int)Math.Ceiling(drawing.Height * scale));
        OfficeRasterImage image = new OfficeRasterImage(width, height, options.Background);
        OfficeRasterCanvas canvas = new OfficeRasterCanvas(image, fonts: drawing.Fonts);
        RenderElements(canvas, drawing.Elements, scale, options.ImageCodec, options.CancellationToken);

        return image;
    }

    private static void RenderElements(
        OfficeRasterCanvas canvas,
        IEnumerable<OfficeDrawingElement> elements,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        foreach (OfficeDrawingElement element in elements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element is OfficeDrawingShape shape) {
                RenderShape(canvas, shape, scale);
            } else if (element is OfficeDrawingText text) {
                RenderText(canvas, text, scale);
            } else if (element is OfficeDrawingRichText richText) {
                RenderRichText(canvas, richText, scale);
            } else if (element is OfficeDrawingImage drawingImage) {
                RenderImage(canvas, drawingImage, scale, imageCodec, cancellationToken);
            } else if (element is OfficeDrawingImagePattern imagePattern) {
                RenderImagePattern(canvas, imagePattern, scale, imageCodec, cancellationToken);
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                RenderTilingPattern(canvas, tilingPattern, scale, imageCodec, cancellationToken);
            } else if (element is OfficeDrawingGroup drawingGroup) {
                RenderGroup(canvas, drawingGroup, scale, imageCodec, cancellationToken);
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                RenderEffectGroup(canvas, effectGroup, scale, imageCodec, cancellationToken);
            }
        }
    }

    /// <summary>
    /// Renders a drawing to PNG bytes.
    /// </summary>
    public static byte[] ToPng(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) =>
        OfficePngWriter.Encode(Render(drawing, scale, background));

    /// <summary>Renders a drawing to PNG bytes with an optional external image codec.</summary>
    public static byte[] ToPng(OfficeDrawing drawing, OfficeDrawingRasterRenderOptions options) => OfficePngWriter.Encode(Render(drawing, options));

    private static void RenderGroup(
        OfficeRasterCanvas canvas,
        OfficeDrawingGroup drawingGroup,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        using (PushGroupClip(canvas, drawingGroup, scale)) {
            var translated = new OfficeDrawing(
                Math.Max(1D, canvas.Width / scale),
                Math.Max(1D, canvas.Height / scale));
            double contentX = drawingGroup.X + drawingGroup.ContentOffsetX;
            double contentY = drawingGroup.Y + drawingGroup.ContentOffsetY;
            if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
                translated.AddDrawingForClippedRendering(drawingGroup.InnerDrawing, contentX, contentY, drawingGroup.FrameTransform.Value);
            } else {
                translated.AddDrawingForClippedRendering(drawingGroup.InnerDrawing, contentX, contentY, null);
            }

            RenderElements(canvas, translated.Elements, scale, imageCodec, cancellationToken);
        }
    }

    private static IDisposable PushGroupClip(OfficeRasterCanvas canvas, OfficeDrawingGroup drawingGroup, double scale) {
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = CreateGroupClipContours(drawingGroup, scale);
        if (contours.Count > 0) {
            return contours.Count == 1 && drawingGroup.ClipPath.Kind != OfficeClipPathKind.Path
                ? canvas.PushClipPolygon(contours[0])
                : PushClipPolygons(canvas, contours, drawingGroup.ClipPath.FillRule);
        }

        if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
            OfficeTransform transform = drawingGroup.FrameTransform.Value.CreateDestinationTransform();
            return canvas.PushClipPolygon(new[] {
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X, drawingGroup.Y)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X + drawingGroup.ClipPath.Width, drawingGroup.Y)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X + drawingGroup.ClipPath.Width, drawingGroup.Y + drawingGroup.ClipPath.Height)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X, drawingGroup.Y + drawingGroup.ClipPath.Height)), scale)
            });
        }

        return canvas.PushClipRectangle(
            drawingGroup.X * scale,
            drawingGroup.Y * scale,
            drawingGroup.ClipPath.Width * scale,
            drawingGroup.ClipPath.Height * scale);
    }

    private static OfficePoint ScalePoint(OfficePoint point, double scale) =>
        new OfficePoint(point.X * scale, point.Y * scale);

    private static void RenderShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        IReadOnlyList<OfficeDrawingShape> glowShapes = CreateGlowShapes(drawingShape);
        for (int i = 0; i < glowShapes.Count; i++) {
            RenderShape(canvas, glowShapes[i], scale);
        }

        IReadOnlyList<OfficeDrawingShape> shadowShapes = CreateShadowShapes(drawingShape);
        for (int i = 0; i < shadowShapes.Count; i++) {
            RenderShape(canvas, shadowShapes[i], scale);
        }

        IDisposable? clipScope = PushShapeClip(canvas, drawingShape, scale);
        try {
            RenderShapeGeometry(canvas, drawingShape, scale);
        } finally {
            clipScope?.Dispose();
        }
    }

    private static void RenderShapeGeometry(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        if (HasNonIdentityTransform(shape.Transform)) {
            RenderTransformedShape(canvas, drawingShape, scale);
            return;
        }

        double x = drawingShape.X * scale;
        double y = drawingShape.Y * scale;
        double width = shape.Width * scale;
        double height = shape.Height * scale;
        OfficeColor? fill = ApplyOpacity(shape.FillColor, shape.FillOpacity);
        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        OfficeRadialGradient? radialGradient = shape.FillRadialGradient == null ? null : ApplyOpacity(shape.FillRadialGradient, shape.FillOpacity);
        OfficeLinearGradient? linearGradient = shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity);
        OfficeRadialGradient? strokeRadialGradient = shape.StrokeRadialGradient == null ? null : ApplyOpacity(shape.StrokeRadialGradient, shape.StrokeOpacity);
        OfficeLinearGradient? strokeLinearGradient = shape.StrokeGradient == null ? null : ApplyOpacity(shape.StrokeGradient, shape.StrokeOpacity);
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
                if (radialGradient != null) {
                    canvas.FillRadialGradientRectangle(x, y, width, height, radialGradient);
                } else if (linearGradient != null) {
                    canvas.FillLinearGradientRectangle(x, y, width, height, linearGradient);
                } else if (fill.HasValue) {
                    canvas.FillRectangle(x, y, width, height, fill.Value);
                }

                if ((stroke.HasValue || strokeLinearGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        DrawGradientOrSolidPolyline(canvas, CreateRectangleContour(x, y, width, height), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
                    } else {
                        DrawGradientOrSolidPolyline(canvas, CreateRectangleContour(x, y, width, height), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
                    }
                }

                break;
            case OfficeShapeKind.RoundedRectangle:
                IReadOnlyList<OfficePoint> rounded = OffsetPoints(CreateRoundedRectangleContour(width, height, shape.CornerRadius * scale, 8), x, y, 1D);
                if (radialGradient != null) {
                    using (canvas.PushClipPolygon(rounded)) {
                        canvas.FillRadialGradientRectangle(x, y, width, height, radialGradient);
                    }
                } else if (linearGradient != null) {
                    using (canvas.PushClipPolygon(rounded)) {
                        canvas.FillLinearGradientRectangle(x, y, width, height, linearGradient);
                    }
                } else if (fill.HasValue) {
                    canvas.FillPolygon(rounded, fill.Value);
                }

                DrawGradientOrSolidPolyline(canvas, rounded, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
                break;
            case OfficeShapeKind.Ellipse:
                if (radialGradient != null) {
                    IReadOnlyList<OfficePoint> ellipse = OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D);
                    using (canvas.PushClipPolygon(ellipse)) {
                        canvas.FillRadialGradientRectangle(x, y, width, height, radialGradient);
                    }
                } else if (linearGradient != null) {
                    IReadOnlyList<OfficePoint> ellipse = OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D);
                    using (canvas.PushClipPolygon(ellipse)) {
                        canvas.FillLinearGradientRectangle(x, y, width, height, linearGradient);
                    }
                } else if (fill.HasValue) {
                    canvas.FillEllipse(x, y, width, height, fill.Value);
                }

                if ((stroke.HasValue || strokeLinearGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        DrawGradientOrSolidPolyline(canvas, OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
                    } else {
                        DrawGradientOrSolidPolyline(canvas, OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
                    }
                }

                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D) RenderLine(canvas, shape, x, y, scale, stroke ?? fill ?? OfficeColor.Black, strokeLinearGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderPolygon(canvas, shape, x, y, scale, fill, linearGradient, radialGradient, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderPath(canvas, shape, x, y, scale, fill, linearGradient, radialGradient, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle);
                break;
        }
    }

    private static void RenderText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale) {
        OfficeTextPadding scaledPadding = text.Padding.Scale(scale);
        double contentX = (text.X * scale) + scaledPadding.Left;
        double contentY = (text.Y * scale) + scaledPadding.Top;
        double contentWidth = (text.Width * scale) - scaledPadding.Horizontal;
        double contentHeight = (text.Height * scale) - scaledPadding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        if (!text.WrapText && !text.ShrinkToFit && !text.HasFrameTransform && text.VerticalAlignment == OfficeTextVerticalAlignment.Top && !text.HasPadding) {
            if (text.TextAdvanceWidth.HasValue) {
                canvas.DrawPositionedText(
                    text.Text,
                    contentX,
                    contentY,
                    contentWidth,
                    contentHeight,
                    text.Color ?? OfficeColor.Black,
                    text.Font.Size * scale,
                    text.Alignment,
                    text.Font.Style,
                    text.Font.FamilyName,
                    text.TextAdvanceWidth.Value * scale);
                return;
            }

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

        double fontSize = Math.Max(1D, text.Font.Size * scale);
        OfficeTextParagraphIndent paragraphIndent = text.ParagraphIndent.Scale(scale);
        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, (text.LineHeight.Value * scale) / fontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D, fontSize);
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
            contentY,
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
            flipVertical: text.FlipVertical);
    }

    private static void RenderRichText(OfficeRasterCanvas canvas, OfficeDrawingRichText text, double scale) {
        OfficeTextPadding scaledPadding = text.Padding.Scale(scale);
        double contentX = (text.X * scale) + scaledPadding.Left;
        double contentY = (text.Y * scale) + scaledPadding.Top;
        double contentWidth = (text.Width * scale) - scaledPadding.Horizontal;
        double contentHeight = (text.Height * scale) - scaledPadding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        IReadOnlyList<OfficeRichTextRun> scaledRuns = ScaleRichTextRuns(text.Runs, scale);
        OfficeTextParagraphIndent paragraphIndent = text.ParagraphIndent.Scale(scale);
        double maxFontSize = 10D * scale;
        for (int i = 0; i < scaledRuns.Count; i++) {
            maxFontSize = Math.Max(maxFontSize, scaledRuns[i].FontSize);
        }

        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, (text.LineHeight.Value * scale) / maxFontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D * scale, maxFontSize);
        Func<string?, double, string?, double> measure = (value, size, family) => canvas.MeasureText(value, size, family);
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            scaledRuns,
            contentWidth,
            contentHeight,
            lineHeightFactor,
            measure,
            text.WrapText,
            text.ShrinkToFit,
            minimumFontSize,
            paragraphIndent);
        OfficeTextBlockRenderer.DrawRasterRichTextBlock(
            canvas,
            layout,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            text.Alignment,
            text.VerticalAlignment,
            text.RotationDegrees,
            text.RotationCenterX * scale,
            text.RotationCenterY * scale,
            flipHorizontal: text.FlipHorizontal,
            flipVertical: text.FlipVertical);
    }

    private static IReadOnlyList<OfficeRichTextRun> ScaleRichTextRuns(IReadOnlyList<OfficeRichTextRun> runs, double scale) {
        var scaled = new List<OfficeRichTextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            scaled.Add(new OfficeRichTextRun(
                run.Text,
                run.FontSize * scale,
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough,
                run.BackgroundColor));
        }

        return scaled;
    }

    private static void RenderImage(
        OfficeRasterCanvas canvas,
        OfficeDrawingImage drawingImage,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        if (TryDecodeImage(
                drawingImage.EncodedBytes,
                drawingImage.ContentType,
                drawingImage.Projection.Width * scale,
                drawingImage.Projection.Height * scale,
                imageCodec,
                cancellationToken,
                out OfficeRasterImage? image) &&
            image != null) {
            if (drawingImage.Opacity < 1D) {
                image = ApplyImageOpacity(image, drawingImage.Opacity);
            }

            canvas.DrawImage(image, drawingImage.Projection.Scale(scale));
        }
    }

    private static bool TryDecodeImage(
        byte[] bytes,
        string? contentType,
        double targetWidth,
        double targetHeight,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken,
        out OfficeRasterImage? image) {
        if (OfficeRasterImageDecoder.TryDecode(bytes, out image) && image != null) return true;
        if (IsSvg(bytes, contentType) &&
            OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? vector, out int unsupportedFeatureCount) &&
            vector != null &&
            unsupportedFeatureCount == 0) {
            cancellationToken.ThrowIfCancellationRequested();
            double scale = ResolveNestedVectorScale(vector, targetWidth, targetHeight);
            image = Render(vector, new OfficeDrawingRasterRenderOptions {
                Scale = scale,
                Background = OfficeColor.Transparent,
                ImageCodec = imageCodec,
                CancellationToken = cancellationToken
            });
            return true;
        }
        return imageCodec != null && imageCodec.TryDecode((byte[])bytes.Clone(), contentType, out image) && image != null;
    }

    private static bool IsSvg(byte[] bytes, string? contentType) =>
        OfficeImageInfo.FromMimeType(contentType) == OfficeImageFormat.Svg ||
        (OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo info) &&
         info.Format == OfficeImageFormat.Svg);

    private static double ResolveNestedVectorScale(
        OfficeDrawing drawing,
        double targetWidth,
        double targetHeight) {
        const long maximumNestedVectorPixels = 16_000_000L;
        double desired = Math.Max(
            Math.Max(1D, targetWidth) / drawing.Width,
            Math.Max(1D, targetHeight) / drawing.Height);
        double safe = Math.Sqrt(
            maximumNestedVectorPixels /
            Math.Max(1D, drawing.Width * drawing.Height));
        return Math.Max(0.000001D, Math.Min(desired, safe));
    }

    private static OfficeRasterImage ApplyImageOpacity(OfficeRasterImage image, double opacity) {
        var result = new OfficeRasterImage(image.Width, image.Height);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                byte alpha = (byte)Math.Round(pixel.A * opacity);
                result.SetPixel(x, y, OfficeColor.FromRgba(pixel.R, pixel.G, pixel.B, alpha));
            }
        }

        return result;
    }

    private static void RenderTransformedShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        OfficeColor? fill = ApplyOpacity(shape.FillColor, shape.FillOpacity);
        OfficeLinearGradient? fillGradient = shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity);
        OfficeRadialGradient? fillRadialGradient = shape.FillRadialGradient == null ? null : ApplyOpacity(shape.FillRadialGradient, shape.FillOpacity);

        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        OfficeLinearGradient? strokeGradient = shape.StrokeGradient == null ? null : ApplyOpacity(shape.StrokeGradient, shape.StrokeOpacity);
        OfficeRadialGradient? strokeRadialGradient = shape.StrokeRadialGradient == null ? null : ApplyOpacity(shape.StrokeRadialGradient, shape.StrokeOpacity);
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.RoundedRectangle:
            case OfficeShapeKind.Ellipse:
                RenderTransformedClosedContour(canvas, drawingShape, scale, CreateShapeContour(shape), fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D) RenderTransformedLine(canvas, drawingShape, scale, stroke ?? fill ?? OfficeColor.Black, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderTransformedClosedContour(canvas, drawingShape, scale, shape.Points, fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderTransformedPath(canvas, drawingShape, scale, fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle);
                break;
        }
    }

    private static void RenderTransformedLine(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor color, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count >= 2) {
            OfficePoint a = TransformShapePoint(drawingShape, shape.Points[0], scale);
            OfficePoint b = TransformShapePoint(drawingShape, shape.Points[1], scale);
            OfficeColor startColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, a, b, a);
            OfficeColor endColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, a, b, b);
            DrawGradientOrSolidPolyline(canvas, new[] { a, b }, color, strokeGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: false, shape.StrokeLineCap);
            RenderLineMarkers(canvas, shape, a, b, startColor, endColor, scale);
        }
    }

    private static void RenderTransformedClosedContour(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, IReadOnlyList<OfficePoint> contour, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        if (contour.Count < 3) {
            return;
        }

        List<OfficePoint> points = TransformShapePoints(drawingShape, contour, scale);
        if (fillGradient != null) {
            fillGradient = TransformShapeFillGradient(drawingShape, scale,
                new[] { (IReadOnlyList<OfficePoint>)points }, fillGradient);
        }
        if (fillRadialGradient != null) canvas.FillRadialGradientPolygon(points, fillRadialGradient);
        else if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        DrawGradientOrSolidPolyline(canvas, points, stroke, strokeGradient, strokeRadialGradient, strokeWidth, drawingShape.Shape.StrokeDashStyle, close: true);
    }

    private static void RenderTransformedPath(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        OfficeShape shape = drawingShape.Shape;
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D);
        if (fillRadialGradient != null || fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(TransformShapePoints(drawingShape, contours[i].Points, scale));
                }
            }

            if (closedContours.Count > 0) {
                if (fillGradient != null) {
                    fillGradient = TransformShapeFillGradient(drawingShape, scale,
                        closedContours, fillGradient);
                }
                if (fillRadialGradient != null || fillGradient != null) {
                    FillGradientPathContours(canvas, closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            for (int i = 0; i < contours.Count; i++) {
                IReadOnlyList<OfficePoint> points = contours[i].Closed
                    ? CloseContour(contours[i].Points)
                    : contours[i].Points;
                DrawGradientOrSolidPolyline(canvas, TransformShapePoints(drawingShape, points, scale), stroke, strokeGradient, strokeRadialGradient, strokeWidth, dashStyle, close: false, shape.StrokeLineCap);
            }

            RenderPathMarkers(canvas, shape, contours, stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient) ?? OfficeColor.Black, scale, strokeGradient, strokeRadialGradient, 0D, 0D, 1D, 1D, point => TransformShapePoint(drawingShape, point, scale));
        }
    }

    private static void RenderLine(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor color, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        if (shape.Points.Count >= 2) {
            OfficePoint a = shape.Points[0];
            OfficePoint b = shape.Points[1];
            OfficePoint start = new OfficePoint(x + (a.X * scale), y + (a.Y * scale));
            OfficePoint end = new OfficePoint(x + (b.X * scale), y + (b.Y * scale));
            OfficeColor startColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, start, end, start);
            OfficeColor endColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, start, end, end);
            DrawGradientOrSolidPolyline(canvas, new[] { start, end }, color, strokeGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: false, shape.StrokeLineCap);
            RenderLineMarkers(canvas, shape, start, end, startColor, endColor, scale);
        }
    }

    private static void RenderLineMarkers(OfficeRasterCanvas canvas, OfficeShape shape, OfficePoint start, OfficePoint end, OfficeColor startColor, OfficeColor endColor, double scale) {
        RenderLineMarker(canvas, shape.StrokeStartMarker, start, new OfficePoint(start.X - end.X, start.Y - end.Y), startColor, scale);
        RenderLineMarker(canvas, shape.StrokeEndMarker, end, new OfficePoint(end.X - start.X, end.Y - start.Y), endColor, scale);
    }

    private static void RenderLineMarker(OfficeRasterCanvas canvas, OfficeLineMarker? marker, OfficePoint tip, OfficePoint lineDirection, OfficeColor color, double scale) {
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(ScaleLineMarker(marker, scale), tip, lineDirection);
        if (contour.Count >= 3) {
            canvas.FillPolygon(contour, color);
        }
    }

    private static void DrawGradientOrSolidPolyline(
        OfficeRasterCanvas canvas,
        IReadOnlyList<OfficePoint> points,
        OfficeColor? stroke,
        OfficeLinearGradient? strokeGradient,
        OfficeRadialGradient? strokeRadialGradient,
        double strokeWidth,
        OfficeStrokeDashStyle dashStyle,
        bool close,
        OfficeStrokeLineCap? lineCap = null) {
        if (strokeWidth <= 0D || points.Count < 2) {
            return;
        }

        if ((strokeGradient == null && strokeRadialGradient == null) || dashStyle != OfficeStrokeDashStyle.Solid) {
            OfficeColor? fallbackStroke = stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient);
            if (!fallbackStroke.HasValue) {
                return;
            }

            if (!close && points.Count == 2 && dashStyle == OfficeStrokeDashStyle.Solid && lineCap.HasValue && lineCap.Value != OfficeStrokeLineCap.Round) {
                DrawCappedLine(canvas, points[0], points[1], fallbackStroke.Value, strokeWidth, lineCap.Value);
            } else if (close) {
                canvas.DrawStyledPolygon(points, fallbackStroke.Value, strokeWidth, dashStyle);
            } else {
                canvas.DrawStyledPolyline(points, fallbackStroke.Value, strokeWidth, dashStyle);
            }

            return;
        }

        IReadOnlyList<OfficePoint> strokePoints = close ? CloseContour(points) : points;
        GetPointBounds(strokePoints, out double x, out double y, out double width, out double height);
        for (int i = 1; i < strokePoints.Count; i++) {
            DrawGradientLineSegment(
                canvas,
                strokePoints[i - 1],
                strokePoints[i],
                strokeGradient,
                strokeRadialGradient,
                x,
                y,
                width,
                height,
                strokeWidth,
                !close && i == 1 ? lineCap : null,
                !close && i == strokePoints.Count - 1 ? lineCap : null);
        }
    }

    private static void DrawCappedLine(OfficeRasterCanvas canvas, OfficePoint start, OfficePoint end, OfficeColor color, double strokeWidth, OfficeStrokeLineCap lineCap) {
        double dx = end.X - start.X;
        double dy = end.Y - start.Y;
        double length = Math.Sqrt((dx * dx) + (dy * dy));
        if (length <= double.Epsilon) {
            return;
        }

        double half = strokeWidth / 2D;
        double ux = dx / length;
        double uy = dy / length;
        double px = -uy * half;
        double py = ux * half;
        double extension = lineCap == OfficeStrokeLineCap.Square ? half : 0D;
        double sx = start.X - (ux * extension);
        double sy = start.Y - (uy * extension);
        double ex = end.X + (ux * extension);
        double ey = end.Y + (uy * extension);
        canvas.FillPolygon(new[] {
            new OfficePoint(sx + px, sy + py),
            new OfficePoint(ex + px, ey + py),
            new OfficePoint(ex - px, ey - py),
            new OfficePoint(sx - px, sy - py)
        }, color);
    }

    private static OfficeColor? GetGradientFallbackStroke(OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient) {
        if (strokeGradient?.Stops.Count > 0) {
            return strokeGradient.Stops[0].Color;
        }

        if (strokeRadialGradient?.Stops.Count > 0) {
            return strokeRadialGradient.Stops[0].Color;
        }

        return null;
    }

    private static OfficeColor SampleLineMarkerColor(OfficeColor fallback, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, OfficePoint start, OfficePoint end, OfficePoint samplePoint) {
        double left = Math.Min(start.X, end.X);
        double top = Math.Min(start.Y, end.Y);
        double width = Math.Abs(end.X - start.X);
        double height = Math.Abs(end.Y - start.Y);
        return SampleStrokeGradient(strokeGradient, strokeRadialGradient, left, top, width, height, samplePoint.X, samplePoint.Y) ?? fallback;
    }

    private static void DrawGradientLineSegment(
        OfficeRasterCanvas canvas,
        OfficePoint start,
        OfficePoint end,
        OfficeLinearGradient? strokeGradient,
        OfficeRadialGradient? strokeRadialGradient,
        double x,
        double y,
        double width,
        double height,
        double strokeWidth,
        OfficeStrokeLineCap? startCap,
        OfficeStrokeLineCap? endCap) {
        double length = Distance(start.X, start.Y, end.X, end.Y);
        if (length <= 0D) {
            return;
        }

        int segments = Math.Max(1, (int)Math.Ceiling(length / 4D));
        for (int segment = 0; segment < segments; segment++) {
            double startRatio = segment / (double)segments;
            double endRatio = (segment + 1) / (double)segments;
            double midRatio = (startRatio + endRatio) / 2D;
            double x1 = start.X + ((end.X - start.X) * startRatio);
            double y1 = start.Y + ((end.Y - start.Y) * startRatio);
            double x2 = start.X + ((end.X - start.X) * endRatio);
            double y2 = start.Y + ((end.Y - start.Y) * endRatio);
            OfficeColor? color = SampleStrokeGradient(strokeGradient, strokeRadialGradient, x, y, width, height, start.X + ((end.X - start.X) * midRatio), start.Y + ((end.Y - start.Y) * midRatio));
            if (color.HasValue) {
                OfficeStrokeLineCap? segmentCap = segment == 0 && startCap.HasValue && startCap.Value != OfficeStrokeLineCap.Round
                    ? startCap
                    : segment == segments - 1 && endCap.HasValue && endCap.Value != OfficeStrokeLineCap.Round
                        ? endCap
                        : null;
                if (segmentCap.HasValue) {
                    DrawCappedLine(canvas, new OfficePoint(x1, y1), new OfficePoint(x2, y2), color.Value, strokeWidth, segmentCap.Value);
                } else {
                    canvas.DrawLine(x1, y1, x2, y2, color.Value, strokeWidth);
                }
            }
        }
    }

    private static OfficeColor? SampleStrokeGradient(OfficeLinearGradient? linearGradient, OfficeRadialGradient? radialGradient, double x, double y, double width, double height, double sampleX, double sampleY) {
        width = Math.Max(width, 0.0001D);
        height = Math.Max(height, 0.0001D);
        double nx = (sampleX - x) / width;
        double ny = (sampleY - y) / height;
        if (radialGradient != null) {
            return InterpolateGradient(radialGradient, ComputeRadialRatio(radialGradient, nx, ny));
        }

        if (linearGradient == null) {
            return null;
        }

        double dx = linearGradient.EndX - linearGradient.StartX;
        double dy = linearGradient.EndY - linearGradient.StartY;
        double lengthSquared = (dx * dx) + (dy * dy);
        if (lengthSquared <= double.Epsilon) {
            return linearGradient.Stops[0].Color;
        }

        double ratio = (((nx - linearGradient.StartX) * dx) + ((ny - linearGradient.StartY) * dy)) / lengthSquared;
        return InterpolateGradient(linearGradient, Clamp(ratio, 0D, 1D));
    }

    private static void GetPointBounds(IReadOnlyList<OfficePoint> points, out double x, out double y, out double width, out double height) {
        double left = points[0].X;
        double top = points[0].Y;
        double right = points[0].X;
        double bottom = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            left = Math.Min(left, points[i].X);
            top = Math.Min(top, points[i].Y);
            right = Math.Max(right, points[i].X);
            bottom = Math.Max(bottom, points[i].Y);
        }

        x = left;
        y = top;
        width = right - left;
        height = bottom - top;
    }

    private static double Distance(double x1, double y1, double x2, double y2) {
        double dx = x2 - x1;
        double dy = y2 - y1;
        return Math.Sqrt((dx * dx) + (dy * dy));
    }

    private static OfficeColor InterpolateGradient(OfficeLinearGradient gradient, double ratio) =>
        InterpolateGradientStops(gradient.Stops, ratio);

    private static OfficeColor InterpolateGradient(OfficeRadialGradient gradient, double ratio) =>
        InterpolateGradientStops(gradient.Stops, ratio);

    private static OfficeColor InterpolateGradientStops(IReadOnlyList<OfficeGradientStop> stops, double ratio) {
        if (ratio <= stops[0].Offset) {
            return stops[0].Color;
        }

        for (int i = 1; i < stops.Count; i++) {
            OfficeGradientStop next = stops[i];
            if (ratio <= next.Offset) {
                OfficeGradientStop previous = stops[i - 1];
                double span = next.Offset - previous.Offset;
                double localRatio = span <= double.Epsilon ? 0D : (ratio - previous.Offset) / span;
                return Interpolate(previous.Color, next.Color, Clamp(localRatio, 0D, 1D));
            }
        }

        return stops[stops.Count - 1].Color;
    }

    private static OfficeColor Interpolate(OfficeColor start, OfficeColor end, double ratio) =>
        OfficeColor.FromRgba(
            InterpolateByte(start.R, end.R, ratio),
            InterpolateByte(start.G, end.G, ratio),
            InterpolateByte(start.B, end.B, ratio),
            InterpolateByte(start.A, end.A, ratio));

    private static byte InterpolateByte(byte start, byte end, double ratio) =>
        (byte)Math.Round(start + ((end - start) * Clamp(ratio, 0D, 1D)));

    private static double ComputeRadialRatio(OfficeRadialGradient gradient, double x, double y) {
        double vx = x - gradient.StartX;
        double vy = y - gradient.StartY;
        double dx = gradient.EndX - gradient.StartX;
        double dy = gradient.EndY - gradient.StartY;
        double dr = gradient.EndRadius - gradient.StartRadius;
        double a = (dx * dx) + (dy * dy) - (dr * dr);
        double b = -2D * ((vx * dx) + (vy * dy) + (gradient.StartRadius * dr));
        double c = (vx * vx) + (vy * vy) - (gradient.StartRadius * gradient.StartRadius);
        if (Math.Abs(a) < 0.0000001D) {
            if (Math.Abs(b) < 0.0000001D) {
                return 0D;
            }

            return Clamp(-c / b, 0D, 1D);
        }

        double discriminant = (b * b) - (4D * a * c);
        if (discriminant < 0D) {
            return 0D;
        }

        double root = Math.Sqrt(discriminant);
        double first = (-b - root) / (2D * a);
        double second = (-b + root) / (2D * a);
        if (first >= 0D && first <= 1D) {
            return first;
        }

        return Clamp(second, 0D, 1D);
    }

    private static double Clamp(double value, double min, double max) =>
        value < min ? min : value > max ? max : value;

    private static OfficeLineMarker? ScaleLineMarker(OfficeLineMarker? marker, double scale) =>
        marker == null ? null : new OfficeLineMarker(marker.Kind, marker.Width * scale, marker.Length * scale);

    private static void RenderPolygon(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        List<OfficePoint> points = OffsetPoints(shape.Points, x, y, scale);
        if (fillRadialGradient != null) canvas.FillRadialGradientPolygon(points, fillRadialGradient);
        else if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        DrawGradientOrSolidPolyline(canvas, points, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape.StrokeDashStyle, close: true);
    }

    private static void RenderPath(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, x, y, scale);
        if (fillRadialGradient != null || fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(contours[i].Points);
                }
            }

            if (closedContours.Count > 0) {
                if (fillRadialGradient != null || fillGradient != null) {
                    FillGradientPathContours(canvas, closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            for (int i = 0; i < contours.Count; i++) {
                IReadOnlyList<OfficePoint> points = contours[i].Closed
                    ? CloseContour(contours[i].Points)
                    : contours[i].Points;
                DrawGradientOrSolidPolyline(canvas, points, stroke, strokeGradient, strokeRadialGradient, strokeWidth, dashStyle, close: false, shape.StrokeLineCap);
            }

            RenderPathMarkers(canvas, shape, contours, stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient) ?? OfficeColor.Black, scale, strokeGradient, strokeRadialGradient, x, y, shape.Width * scale, shape.Height * scale);
        }
    }

    private static void RenderPathMarkers(OfficeRasterCanvas canvas, OfficeShape shape, IReadOnlyList<OfficeFlattenedPathContour> contours, OfficeColor fallbackColor, double scale, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double gradientX, double gradientY, double gradientWidth, double gradientHeight, Func<OfficePoint, OfficePoint>? transformPoint = null) {
        if (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null) {
            return;
        }

        OfficeFlattenedPathContour? firstOpen = null;
        OfficeFlattenedPathContour? lastOpen = null;
        for (int i = 0; i < contours.Count; i++) {
            if (!contours[i].Closed && contours[i].Points.Count >= 2) {
                firstOpen ??= contours[i];
                lastOpen = contours[i];
            }
        }

        if (firstOpen != null) {
            OfficePoint start = TransformMarkerPoint(firstOpen.Points[0], transformPoint);
            OfficePoint next = TransformMarkerPoint(firstOpen.Points[1], transformPoint);
            OfficeColor startColor = SampleStrokeGradient(strokeGradient, strokeRadialGradient, gradientX, gradientY, gradientWidth, gradientHeight, firstOpen.Points[0].X, firstOpen.Points[0].Y) ?? fallbackColor;
            RenderLineMarker(canvas, shape.StrokeStartMarker, start, new OfficePoint(start.X - next.X, start.Y - next.Y), startColor, scale);
        }

        if (lastOpen != null) {
            IReadOnlyList<OfficePoint> points = lastOpen.Points;
            OfficePoint end = TransformMarkerPoint(points[points.Count - 1], transformPoint);
            OfficePoint previous = TransformMarkerPoint(points[points.Count - 2], transformPoint);
            OfficeColor endColor = SampleStrokeGradient(strokeGradient, strokeRadialGradient, gradientX, gradientY, gradientWidth, gradientHeight, points[points.Count - 1].X, points[points.Count - 1].Y) ?? fallbackColor;
            RenderLineMarker(canvas, shape.StrokeEndMarker, end, new OfficePoint(end.X - previous.X, end.Y - previous.Y), endColor, scale);
        }
    }

    private static OfficePoint TransformMarkerPoint(OfficePoint point, Func<OfficePoint, OfficePoint>? transformPoint) =>
        transformPoint == null ? point : transformPoint(point);

    private static IReadOnlyList<OfficePoint> CloseContour(IReadOnlyList<OfficePoint> points) {
        if (points.Count < 2) {
            return points;
        }

        var closed = new List<OfficePoint>(points.Count + 1);
        for (int i = 0; i < points.Count; i++) {
            closed.Add(points[i]);
        }

        closed.Add(points[0]);
        return closed;
    }

    private static IReadOnlyList<OfficePoint> CreateRectangleContour(double x, double y, double width, double height) =>
        new[] {
            new OfficePoint(x, y),
            new OfficePoint(x + width, y),
            new OfficePoint(x + width, y + height),
            new OfficePoint(x, y + height)
        };

    private static IDisposable? PushShapeClip(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeClipPath? clipPath = drawingShape.Shape.ClipPath;
        if (clipPath == null) {
            return null;
        }

        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = CreateClipContours(drawingShape, clipPath, scale);
        if (contours.Count == 0) {
            return null;
        }

        return contours.Count == 1 && clipPath.Kind != OfficeClipPathKind.Path
            ? canvas.PushClipPolygon(contours[0])
            : PushClipPolygons(canvas, contours, clipPath.FillRule);
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateClipContours(OfficeDrawingShape drawingShape, OfficeClipPath clipPath, double scale) {
        return CreateClipContours(clipPath, contour => TransformClipContour(drawingShape, contour, scale));
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateGroupClipContours(OfficeDrawingGroup drawingGroup, double scale) {
        OfficeTransform? transform = drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform
            ? drawingGroup.FrameTransform.Value.CreateDestinationTransform()
            : null;
        return CreateClipContours(
            drawingGroup.ClipPath,
            contour => TransformGroupClipContour(drawingGroup, contour, scale, transform));
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateClipContours(OfficeClipPath clipPath, Func<IReadOnlyList<OfficePoint>, IReadOnlyList<OfficePoint>> transformContour) {
        IReadOnlyList<OfficePoint> contour;
        switch (clipPath.Kind) {
            case OfficeClipPathKind.Rectangle:
                contour = new[] {
                    new OfficePoint(0D, 0D),
                    new OfficePoint(clipPath.Width, 0D),
                    new OfficePoint(clipPath.Width, clipPath.Height),
                    new OfficePoint(0D, clipPath.Height)
                };
                return new[] { transformContour(contour) };
            case OfficeClipPathKind.RoundedRectangle:
                contour = CreateRoundedRectangleContour(clipPath.Width, clipPath.Height, clipPath.CornerRadius, 8);
                return new[] { transformContour(contour) };
            case OfficeClipPathKind.Path:
                IReadOnlyList<OfficeFlattenedPathContour> flattened = OfficePathFlattener.Flatten(clipPath.Commands, 0D, 0D, 1D);
                List<IReadOnlyList<OfficePoint>> contours = new List<IReadOnlyList<OfficePoint>>();
                for (int i = 0; i < flattened.Count; i++) {
                    if (flattened[i].Closed && flattened[i].Points.Count >= 3) {
                        contours.Add(transformContour(flattened[i].Points));
                    }
                }

                return contours;
            default:
                return Array.Empty<IReadOnlyList<OfficePoint>>();
        }
    }

    private static void FillPathContours(OfficeRasterCanvas canvas, IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color, OfficeFillRule fillRule) {
        if (fillRule == OfficeFillRule.NonZero) {
            canvas.FillPolygonsNonZero(contours, color);
        } else {
            canvas.FillPolygonsEvenOdd(contours, color);
        }
    }

    private static void FillGradientPathContours(OfficeRasterCanvas canvas, IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeLinearGradient? linearGradient, OfficeRadialGradient? radialGradient, OfficeFillRule fillRule) {
        if (contours.Count == 1 && fillRule != OfficeFillRule.NonZero) {
            if (radialGradient != null) {
                canvas.FillRadialGradientPolygon(contours[0], radialGradient);
            } else if (linearGradient != null) {
                canvas.FillLinearGradientPolygon(contours[0], linearGradient);
            }

            return;
        }

        if (!TryGetContourBounds(contours, out double left, out double top, out double right, out double bottom)) {
            return;
        }

        var bounds = new[] {
            new OfficePoint(left, top),
            new OfficePoint(right, top),
            new OfficePoint(right, bottom),
            new OfficePoint(left, bottom)
        };
        using (PushClipPolygons(canvas, contours, fillRule)) {
            if (radialGradient != null) {
                canvas.FillRadialGradientPolygon(bounds, radialGradient);
            } else if (linearGradient != null) {
                canvas.FillLinearGradientPolygon(bounds, linearGradient);
            }
        }
    }

    private static bool TryGetContourBounds(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, out double left, out double top, out double right, out double bottom) {
        left = 0D;
        top = 0D;
        right = 0D;
        bottom = 0D;
        bool hasPoint = false;
        for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
            IReadOnlyList<OfficePoint> contour = contours[contourIndex];
            for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                OfficePoint point = contour[pointIndex];
                if (!hasPoint) {
                    left = right = point.X;
                    top = bottom = point.Y;
                    hasPoint = true;
                    continue;
                }

                if (point.X < left) left = point.X;
                if (point.Y < top) top = point.Y;
                if (point.X > right) right = point.X;
                if (point.Y > bottom) bottom = point.Y;
            }
        }

        return hasPoint && right > left && bottom > top;
    }

    private static OfficeLinearGradient TransformShapeFillGradient(
        OfficeDrawingShape drawingShape,
        double scale,
        IReadOnlyList<IReadOnlyList<OfficePoint>> transformedContours,
        OfficeLinearGradient gradient) {
        if (!TryGetContourBounds(transformedContours, out double left, out double top,
                out double right, out double bottom)) {
            return gradient;
        }

        OfficeShape shape = drawingShape.Shape;
        OfficePoint start = TransformShapePoint(drawingShape, new OfficePoint(
            gradient.StartX * shape.Width,
            gradient.StartY * shape.Height), scale);
        OfficePoint end = TransformShapePoint(drawingShape, new OfficePoint(
            gradient.EndX * shape.Width,
            gradient.EndY * shape.Height), scale);
        double width = right - left;
        double height = bottom - top;
        double startX = (start.X - left) / width;
        double startY = (start.Y - top) / height;
        double endX = (end.X - left) / width;
        double endY = (end.Y - top) / height;
        if (startX.Equals(endX) && startY.Equals(endY)) {
            return gradient;
        }

        return OfficeLinearGradient.CreateImported(startX, startY, endX, endY,
            gradient.Stops);
    }

    private static IDisposable PushClipPolygons(OfficeRasterCanvas canvas, IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule) =>
        fillRule == OfficeFillRule.NonZero
            ? canvas.PushClipPolygonsNonZero(contours)
            : canvas.PushClipPolygonsEvenOdd(contours);

    private static IReadOnlyList<OfficePoint> TransformClipContour(OfficeDrawingShape drawingShape, IReadOnlyList<OfficePoint> contour, double scale) =>
        HasNonIdentityTransform(drawingShape.Shape.Transform)
            ? TransformShapePoints(drawingShape, contour, scale)
            : OffsetPoints(contour, drawingShape.X * scale, drawingShape.Y * scale, scale);

    private static IReadOnlyList<OfficePoint> TransformGroupClipContour(OfficeDrawingGroup drawingGroup, IReadOnlyList<OfficePoint> contour, double scale, OfficeTransform? transform) {
        List<OfficePoint> points = new List<OfficePoint>(contour.Count);
        for (int i = 0; i < contour.Count; i++) {
            OfficePoint point = new OfficePoint(drawingGroup.X + contour[i].X, drawingGroup.Y + contour[i].Y);
            if (transform.HasValue) {
                point = transform.Value.TransformPoint(point);
            }

            points.Add(ScalePoint(point, scale));
        }

        return points;
    }

    private static List<OfficePoint> OffsetPoints(IReadOnlyList<OfficePoint> source, double x, double y, double scale) {
        List<OfficePoint> points = new List<OfficePoint>(source.Count);
        for (int i = 0; i < source.Count; i++) {
            points.Add(new OfficePoint(x + (source[i].X * scale), y + (source[i].Y * scale)));
        }

        return points;
    }

    private static IReadOnlyList<OfficePoint> CreateShapeContour(OfficeShape shape) {
        switch (shape.Kind) {
            case OfficeShapeKind.Ellipse:
                return CreateEllipseContour(shape.Width, shape.Height, 72);
            case OfficeShapeKind.RoundedRectangle:
                return CreateRoundedRectangleContour(shape.Width, shape.Height, shape.CornerRadius, 8);
            case OfficeShapeKind.Rectangle:
            default:
                return new[] {
                    new OfficePoint(0D, 0D),
                    new OfficePoint(shape.Width, 0D),
                    new OfficePoint(shape.Width, shape.Height),
                    new OfficePoint(0D, shape.Height)
                };
        }
    }

    private static IReadOnlyList<OfficePoint> CreateEllipseContour(double width, double height, int segments) {
        List<OfficePoint> points = new List<OfficePoint>(segments);
        double centerX = width / 2D;
        double centerY = height / 2D;
        for (int i = 0; i < segments; i++) {
            double radians = (Math.PI * 2D * i) / segments;
            points.Add(new OfficePoint(
                centerX + (Math.Cos(radians) * centerX),
                centerY + (Math.Sin(radians) * centerY)));
        }

        return points;
    }

    private static IReadOnlyList<OfficePoint> CreateRoundedRectangleContour(double width, double height, double radius, int cornerSegments) {
        double r = Math.Max(0D, Math.Min(radius, Math.Min(width, height) / 2D));
        if (r <= 0D) {
            return CreateShapeContour(OfficeShape.Rectangle(width, height));
        }

        List<OfficePoint> points = new List<OfficePoint>(cornerSegments * 4);
        AddArc(points, width - r, r, r, -90D, 0D, cornerSegments);
        AddArc(points, width - r, height - r, r, 0D, 90D, cornerSegments);
        AddArc(points, r, height - r, r, 90D, 180D, cornerSegments);
        AddArc(points, r, r, r, 180D, 270D, cornerSegments);
        return points;
    }

    private static void AddArc(List<OfficePoint> points, double centerX, double centerY, double radius, double startDegrees, double endDegrees, int segments) {
        for (int i = 0; i <= segments; i++) {
            double degrees = startDegrees + ((endDegrees - startDegrees) * i / segments);
            double radians = OfficeGeometry.DegreesToRadians(degrees);
            points.Add(new OfficePoint(
                centerX + (Math.Cos(radians) * radius),
                centerY + (Math.Sin(radians) * radius)));
        }
    }

    private static List<OfficePoint> TransformShapePoints(OfficeDrawingShape drawingShape, IReadOnlyList<OfficePoint> points, double scale) {
        List<OfficePoint> transformed = new List<OfficePoint>(points.Count);
        for (int i = 0; i < points.Count; i++) {
            transformed.Add(TransformShapePoint(drawingShape, points[i], scale));
        }

        return transformed;
    }

    private static OfficePoint TransformShapePoint(OfficeDrawingShape drawingShape, OfficePoint point, double scale) {
        OfficePoint local = drawingShape.Shape.Transform.HasValue
            ? drawingShape.Shape.Transform.Value.TransformPoint(point)
            : point;
        return new OfficePoint((drawingShape.X + local.X) * scale, (drawingShape.Y + local.Y) * scale);
    }

    private static bool HasNonIdentityTransform(OfficeTransform? transform) =>
        transform.HasValue && transform.Value != OfficeTransform.Identity;

    private static OfficeColor? ApplyOpacity(OfficeColor? color, double? opacity) {
        if (!color.HasValue) return null;
        if (!opacity.HasValue) return color;
        double clamped = opacity.Value < 0D ? 0D : opacity.Value > 1D ? 1D : opacity.Value;
        return OfficeColor.FromRgba(color.Value.R, color.Value.G, color.Value.B, (byte)Math.Round(color.Value.A * clamped));
    }

    private static OfficeColor? ApplyOpacity(OfficeColor color, double? opacity) =>
        ApplyOpacity((OfficeColor?)color, opacity);

    private static OfficeLinearGradient ApplyOpacity(OfficeLinearGradient gradient, double? opacity) {
        if (!opacity.HasValue) {
            return gradient;
        }

        var stops = new List<OfficeGradientStop>(gradient.Stops.Count);
        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            stops.Add(new OfficeGradientStop(
                stop.Offset,
                ApplyOpacity(stop.Color, opacity) ?? stop.Color));
        }

        return OfficeLinearGradient.CreateImported(
            gradient.StartX,
            gradient.StartY,
            gradient.EndX,
            gradient.EndY,
            stops);
    }

    private static OfficeRadialGradient ApplyOpacity(OfficeRadialGradient gradient, double? opacity) {
        if (!opacity.HasValue) {
            return gradient;
        }

        var stops = new List<OfficeGradientStop>(gradient.Stops.Count);
        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            stops.Add(new OfficeGradientStop(
                stop.Offset,
                ApplyOpacity(stop.Color, opacity) ?? stop.Color));
        }

        return new OfficeRadialGradient(
            gradient.StartX,
            gradient.StartY,
            gradient.StartRadiusX,
            gradient.StartRadiusY,
            gradient.EndX,
            gradient.EndY,
            gradient.EndRadiusX,
            gradient.EndRadiusY,
            stops);
    }

    private static IReadOnlyList<OfficeDrawingShape> CreateGlowShapes(OfficeDrawingShape drawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeGlow? glow = shape.Glow;
        if (glow == null || glow.Radius <= 0D || glow.Opacity <= 0D || glow.Color.A == 0) {
            return Array.Empty<OfficeDrawingShape>();
        }

        const int layers = 4;
        var glowShapes = new List<OfficeDrawingShape>(layers);
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        for (int i = layers; i >= 1; i--) {
            double factor = i / (double)layers;
            OfficeShape glowShape = shape.Clone();
            glowShape.Shadow = null;
            glowShape.Glow = null;
            glowShape.FillColor = null;
            glowShape.FillGradient = null;
            glowShape.FillRadialGradient = null;
            glowShape.StrokeColor = glow.Color;
            glowShape.StrokeGradient = null;
            glowShape.StrokeRadialGradient = null;
            glowShape.StrokeWidth = Math.Max(1D, baseStrokeWidth + glow.Radius * 2D * factor);
            glowShape.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
            glowShape.StrokeStartMarker = null;
            glowShape.StrokeEndMarker = null;
            glowShape.StrokeOpacity = ComputeGlowLayerOpacity(glow.Opacity, layers - i + 1);
            glowShapes.Add(new OfficeDrawingShape(glowShape, drawingShape.X, drawingShape.Y));
        }

        return glowShapes;
    }

    private static double ComputeGlowLayerOpacity(double opacity, int layerDepth) {
        double clamped = opacity < 0D ? 0D : opacity > 1D ? 1D : opacity;
        return 1D - Math.Pow(1D - clamped, layerDepth + 1);
    }

    private static IReadOnlyList<OfficeDrawingShape> CreateShadowShapes(OfficeDrawingShape drawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) {
            return Array.Empty<OfficeDrawingShape>();
        }

        bool hasStroke = shape.Kind == OfficeShapeKind.Line ||
            (shape.StrokeWidth > 0D &&
                (shape.StrokeRadialGradient != null ||
                 shape.StrokeGradient != null ||
                 (shape.StrokeColor.HasValue && shape.StrokeColor.Value.A > 0)));
        bool hasFill = shape.Kind != OfficeShapeKind.Line &&
            (shape.FillRadialGradient != null || shape.FillGradient != null || (shape.FillColor.HasValue && shape.FillColor.Value.A > 0));
        OfficeDrawingShape coreShadow = CreateShadowShape(drawingShape, shadow, hasStroke, hasFill, Math.Max(0D, shape.StrokeWidth), shadow.Opacity);
        if (shadow.BlurRadius <= 0D) {
            return new[] { coreShadow };
        }

        const int layers = 4;
        var shadowShapes = new List<OfficeDrawingShape>(layers + 1);
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        for (int i = layers; i >= 1; i--) {
            double factor = i / (double)layers;
            double opacity = shadow.Opacity * (0.04D + (layers - i + 1) * 0.05D);
            shadowShapes.Add(CreateShadowShape(
                drawingShape,
                shadow,
                hasStroke: true,
                hasFill: hasFill,
                strokeWidth: Math.Max(1D, baseStrokeWidth + shadow.BlurRadius * 2D * factor),
                opacity: opacity));
        }

        shadowShapes.Add(coreShadow);
        return shadowShapes;
    }

    private static OfficeDrawingShape CreateShadowShape(OfficeDrawingShape drawingShape, OfficeShadow shadow, bool hasStroke, bool hasFill, double strokeWidth, double opacity) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShape shadowShape = shape.Clone();
        shadowShape.Shadow = null;
        shadowShape.Glow = null;
        shadowShape.FillGradient = null;
        shadowShape.FillRadialGradient = null;
        shadowShape.FillColor = hasFill || !hasStroke ? shadow.Color : null;
        shadowShape.FillOpacity = opacity;
        shadowShape.StrokeColor = hasStroke ? shadow.Color : null;
        shadowShape.StrokeGradient = null;
        shadowShape.StrokeRadialGradient = null;
        shadowShape.StrokeWidth = strokeWidth;
        shadowShape.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
        shadowShape.StrokeStartMarker = null;
        shadowShape.StrokeEndMarker = null;
        shadowShape.StrokeOpacity = opacity;

        return CreateOffsetEffectShape(shadowShape, drawingShape.X + shadow.OffsetX, drawingShape.Y + shadow.OffsetY);
    }

    private static OfficeDrawingShape CreateOffsetEffectShape(OfficeShape shape, double x, double y) {
        double clampedX = Math.Max(0D, x);
        double clampedY = Math.Max(0D, y);
        double offsetX = x - clampedX;
        double offsetY = y - clampedY;
        if (offsetX != 0D || offsetY != 0D) {
            shape = shape.Clone();
            OfficeTransform offsetTransform = OfficeTransform.Translate(offsetX, offsetY);
            shape.Transform = shape.Transform.HasValue ? offsetTransform.Then(shape.Transform.Value) : offsetTransform;
        }

        return new OfficeDrawingShape(shape, clampedX, clampedY);
    }
}
