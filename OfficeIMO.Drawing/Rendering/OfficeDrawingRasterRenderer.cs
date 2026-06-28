using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free raster renderer for <see cref="OfficeDrawing"/> scenes.
/// </summary>
public static class OfficeDrawingRasterRenderer {
    /// <summary>
    /// Renders a drawing to an RGBA raster image.
    /// </summary>
    public static OfficeRasterImage Render(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Scale must be a finite positive number.");
        }

        int width = Math.Max(1, (int)Math.Ceiling(drawing.Width * scale));
        int height = Math.Max(1, (int)Math.Ceiling(drawing.Height * scale));
        OfficeRasterImage image = new OfficeRasterImage(width, height, background);
        OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
        RenderElements(canvas, drawing.Elements, scale);

        return image;
    }

    private static void RenderElements(OfficeRasterCanvas canvas, IEnumerable<OfficeDrawingElement> elements, double scale) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingShape shape) {
                RenderShape(canvas, shape, scale);
            } else if (element is OfficeDrawingText text) {
                RenderText(canvas, text, scale);
            } else if (element is OfficeDrawingRichText richText) {
                RenderRichText(canvas, richText, scale);
            } else if (element is OfficeDrawingImage drawingImage) {
                RenderImage(canvas, drawingImage, scale);
            } else if (element is OfficeDrawingGroup drawingGroup) {
                RenderGroup(canvas, drawingGroup, scale);
            }
        }
    }

    /// <summary>
    /// Renders a drawing to PNG bytes.
    /// </summary>
    public static byte[] ToPng(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) =>
        OfficePngWriter.Encode(Render(drawing, scale, background));

    private static void RenderGroup(OfficeRasterCanvas canvas, OfficeDrawingGroup drawingGroup, double scale) {
        using (PushGroupClip(canvas, drawingGroup, scale)) {
            var translated = new OfficeDrawing(
                Math.Max(1D, drawingGroup.X + drawingGroup.InnerDrawing.Width),
                Math.Max(1D, drawingGroup.Y + drawingGroup.InnerDrawing.Height));
            if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
                translated.AddDrawing(drawingGroup.InnerDrawing, drawingGroup.X, drawingGroup.Y, drawingGroup.FrameTransform.Value);
            } else {
                translated.AddDrawing(drawingGroup.InnerDrawing, drawingGroup.X, drawingGroup.Y);
            }

            RenderElements(canvas, translated.Elements, scale);
        }
    }

    private static IDisposable PushGroupClip(OfficeRasterCanvas canvas, OfficeDrawingGroup drawingGroup, double scale) {
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
        if (TryCreateShadowShape(drawingShape, out OfficeDrawingShape shadowShape)) {
            RenderShape(canvas, shadowShape, scale);
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
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
                if (shape.FillGradient != null) {
                    canvas.FillLinearGradientRectangle(x, y, width, height, ApplyOpacity(shape.FillGradient, shape.FillOpacity));
                } else if (fill.HasValue) {
                    canvas.FillRectangle(x, y, width, height, fill.Value);
                }

                if (stroke.HasValue && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        canvas.DrawRectangle(x, y, width, height, stroke.Value, strokeWidth);
                    } else {
                        canvas.DrawStyledPolygon(CreateRectangleContour(x, y, width, height), stroke.Value, strokeWidth, shape.StrokeDashStyle);
                    }
                }

                break;
            case OfficeShapeKind.RoundedRectangle:
                IReadOnlyList<OfficePoint> rounded = OffsetPoints(CreateRoundedRectangleContour(width, height, shape.CornerRadius * scale, 8), x, y, 1D);
                if (shape.FillGradient != null) {
                    using (canvas.PushClipPolygon(rounded)) {
                        canvas.FillLinearGradientRectangle(x, y, width, height, ApplyOpacity(shape.FillGradient, shape.FillOpacity));
                    }
                } else if (fill.HasValue) {
                    canvas.FillPolygon(rounded, fill.Value);
                }

                if (stroke.HasValue && strokeWidth > 0D) canvas.DrawStyledPolygon(rounded, stroke.Value, strokeWidth, shape.StrokeDashStyle);
                break;
            case OfficeShapeKind.Ellipse:
                if (shape.FillGradient != null) {
                    IReadOnlyList<OfficePoint> ellipse = OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D);
                    using (canvas.PushClipPolygon(ellipse)) {
                        canvas.FillLinearGradientRectangle(x, y, width, height, ApplyOpacity(shape.FillGradient, shape.FillOpacity));
                    }
                } else if (fill.HasValue) {
                    canvas.FillEllipse(x, y, width, height, fill.Value);
                }

                if (stroke.HasValue && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        canvas.DrawEllipse(x, y, width, height, stroke.Value, strokeWidth);
                    } else {
                        canvas.DrawStyledPolygon(OffsetPoints(CreateEllipseContour(width, height, 96), x, y, 1D), stroke.Value, strokeWidth, shape.StrokeDashStyle);
                    }
                }

                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D) RenderLine(canvas, shape, x, y, scale, stroke ?? fill ?? OfficeColor.Black, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderPolygon(canvas, shape, x, y, scale, fill, shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity), stroke, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderPath(canvas, shape, x, y, scale, fill, shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity), stroke, strokeWidth, shape.StrokeDashStyle);
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
                shrinkToFit: text.ShrinkToFit,
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

    private static void RenderImage(OfficeRasterCanvas canvas, OfficeDrawingImage drawingImage, double scale) {
        if (OfficeRasterImageDecoder.TryDecode(drawingImage.Bytes, out OfficeRasterImage? image) && image != null) {
            canvas.DrawImage(image, drawingImage.Projection.Scale(scale));
        }
    }

    private static void RenderTransformedShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        OfficeColor? fill = ApplyOpacity(shape.FillColor, shape.FillOpacity);
        OfficeLinearGradient? fillGradient = shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity);

        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.RoundedRectangle:
            case OfficeShapeKind.Ellipse:
                RenderTransformedClosedContour(canvas, drawingShape, scale, CreateShapeContour(shape), fill, fillGradient, stroke, strokeWidth);
                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D) RenderTransformedLine(canvas, drawingShape, scale, stroke ?? fill ?? OfficeColor.Black, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderTransformedClosedContour(canvas, drawingShape, scale, shape.Points, fill, fillGradient, stroke, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderTransformedPath(canvas, drawingShape, scale, fill, fillGradient, stroke, strokeWidth, shape.StrokeDashStyle);
                break;
        }
    }

    private static void RenderTransformedLine(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor color, double strokeWidth) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count >= 2) {
            OfficePoint a = TransformShapePoint(drawingShape, shape.Points[0], scale);
            OfficePoint b = TransformShapePoint(drawingShape, shape.Points[1], scale);
            canvas.DrawStyledLine(a.X, a.Y, b.X, b.Y, color, strokeWidth, shape.StrokeDashStyle);
            RenderLineMarkers(canvas, shape, a, b, color, scale);
        }
    }

    private static void RenderTransformedClosedContour(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, IReadOnlyList<OfficePoint> contour, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeColor? stroke, double strokeWidth) {
        if (contour.Count < 3) {
            return;
        }

        List<OfficePoint> points = TransformShapePoints(drawingShape, contour, scale);
        if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        if (stroke.HasValue && strokeWidth > 0D) canvas.DrawStyledPolygon(points, stroke.Value, strokeWidth, drawingShape.Shape.StrokeDashStyle);
    }

    private static void RenderTransformedPath(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeColor? stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        OfficeShape shape = drawingShape.Shape;
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D);
        if (fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(TransformShapePoints(drawingShape, contours[i].Points, scale));
                }
            }

            if (closedContours.Count > 0) {
                if (fillGradient != null) {
                    for (int i = 0; i < closedContours.Count; i++) {
                        canvas.FillLinearGradientPolygon(closedContours[i], fillGradient);
                    }
                } else {
                    canvas.FillPolygonsEvenOdd(closedContours, fill!.Value);
                }
            }
        }

        if (stroke.HasValue && strokeWidth > 0D) {
            for (int i = 0; i < contours.Count; i++) {
                IReadOnlyList<OfficePoint> points = contours[i].Closed
                    ? CloseContour(contours[i].Points)
                    : contours[i].Points;
                canvas.DrawStyledPolyline(TransformShapePoints(drawingShape, points, scale), stroke.Value, strokeWidth, dashStyle);
            }
        }
    }

    private static void RenderLine(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor color, double strokeWidth) {
        if (shape.Points.Count >= 2) {
            OfficePoint a = shape.Points[0];
            OfficePoint b = shape.Points[1];
            OfficePoint start = new OfficePoint(x + (a.X * scale), y + (a.Y * scale));
            OfficePoint end = new OfficePoint(x + (b.X * scale), y + (b.Y * scale));
            canvas.DrawStyledLine(start.X, start.Y, end.X, end.Y, color, strokeWidth, shape.StrokeDashStyle);
            RenderLineMarkers(canvas, shape, start, end, color, scale);
        }
    }

    private static void RenderLineMarkers(OfficeRasterCanvas canvas, OfficeShape shape, OfficePoint start, OfficePoint end, OfficeColor color, double scale) {
        RenderLineMarker(canvas, shape.StrokeStartMarker, start, new OfficePoint(start.X - end.X, start.Y - end.Y), color, scale);
        RenderLineMarker(canvas, shape.StrokeEndMarker, end, new OfficePoint(end.X - start.X, end.Y - start.Y), color, scale);
    }

    private static void RenderLineMarker(OfficeRasterCanvas canvas, OfficeLineMarker? marker, OfficePoint tip, OfficePoint lineDirection, OfficeColor color, double scale) {
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(ScaleLineMarker(marker, scale), tip, lineDirection);
        if (contour.Count >= 3) {
            canvas.FillPolygon(contour, color);
        }
    }

    private static OfficeLineMarker? ScaleLineMarker(OfficeLineMarker? marker, double scale) =>
        marker == null ? null : new OfficeLineMarker(marker.Kind, marker.Width * scale, marker.Length * scale);

    private static void RenderPolygon(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeColor? stroke, double strokeWidth) {
        List<OfficePoint> points = OffsetPoints(shape.Points, x, y, scale);
        if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        if (stroke.HasValue && strokeWidth > 0D) canvas.DrawStyledPolygon(points, stroke.Value, strokeWidth, shape.StrokeDashStyle);
    }

    private static void RenderPath(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeColor? stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, x, y, scale);
        if (fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(contours[i].Points);
                }
            }

            if (closedContours.Count > 0) {
                if (fillGradient != null) {
                    for (int i = 0; i < closedContours.Count; i++) {
                        canvas.FillLinearGradientPolygon(closedContours[i], fillGradient);
                    }
                } else {
                    canvas.FillPolygonsEvenOdd(closedContours, fill!.Value);
                }
            }
        }

        if (stroke.HasValue && strokeWidth > 0D) {
            for (int i = 0; i < contours.Count; i++) {
                IReadOnlyList<OfficePoint> points = contours[i].Closed
                    ? CloseContour(contours[i].Points)
                    : contours[i].Points;
                canvas.DrawStyledPolyline(points, stroke.Value, strokeWidth, dashStyle);
            }

            RenderPathMarkers(canvas, shape, contours, stroke.Value, scale);
        }
    }

    private static void RenderPathMarkers(OfficeRasterCanvas canvas, OfficeShape shape, IReadOnlyList<OfficeFlattenedPathContour> contours, OfficeColor color, double scale) {
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
            OfficePoint start = firstOpen.Points[0];
            OfficePoint next = firstOpen.Points[1];
            RenderLineMarker(canvas, shape.StrokeStartMarker, start, new OfficePoint(start.X - next.X, start.Y - next.Y), color, scale);
        }

        if (lastOpen != null) {
            IReadOnlyList<OfficePoint> points = lastOpen.Points;
            OfficePoint end = points[points.Count - 1];
            OfficePoint previous = points[points.Count - 2];
            RenderLineMarker(canvas, shape.StrokeEndMarker, end, new OfficePoint(end.X - previous.X, end.Y - previous.Y), color, scale);
        }
    }

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

        return contours.Count == 1
            ? canvas.PushClipPolygon(contours[0])
            : canvas.PushClipPolygonsEvenOdd(contours);
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateClipContours(OfficeDrawingShape drawingShape, OfficeClipPath clipPath, double scale) {
        IReadOnlyList<OfficePoint> contour;
        switch (clipPath.Kind) {
            case OfficeClipPathKind.Rectangle:
                contour = new[] {
                    new OfficePoint(0D, 0D),
                    new OfficePoint(clipPath.Width, 0D),
                    new OfficePoint(clipPath.Width, clipPath.Height),
                    new OfficePoint(0D, clipPath.Height)
                };
                return new[] { TransformClipContour(drawingShape, contour, scale) };
            case OfficeClipPathKind.RoundedRectangle:
                contour = CreateRoundedRectangleContour(clipPath.Width, clipPath.Height, clipPath.CornerRadius, 8);
                return new[] { TransformClipContour(drawingShape, contour, scale) };
            case OfficeClipPathKind.Path:
                IReadOnlyList<OfficeFlattenedPathContour> flattened = OfficePathFlattener.Flatten(clipPath.Commands, 0D, 0D, 1D);
                List<IReadOnlyList<OfficePoint>> contours = new List<IReadOnlyList<OfficePoint>>();
                for (int i = 0; i < flattened.Count; i++) {
                    if (flattened[i].Closed && flattened[i].Points.Count >= 3) {
                        contours.Add(TransformClipContour(drawingShape, flattened[i].Points, scale));
                    }
                }

                return contours;
            default:
                return Array.Empty<IReadOnlyList<OfficePoint>>();
        }
    }

    private static IReadOnlyList<OfficePoint> TransformClipContour(OfficeDrawingShape drawingShape, IReadOnlyList<OfficePoint> contour, double scale) =>
        HasNonIdentityTransform(drawingShape.Shape.Transform)
            ? TransformShapePoints(drawingShape, contour, scale)
            : OffsetPoints(contour, drawingShape.X * scale, drawingShape.Y * scale, scale);

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

        return new OfficeLinearGradient(
            gradient.StartX,
            gradient.StartY,
            gradient.EndX,
            gradient.EndY,
            new OfficeGradientStop(gradient.Stops[0].Offset, ApplyOpacity(gradient.Stops[0].Color, opacity) ?? gradient.Stops[0].Color),
            new OfficeGradientStop(gradient.Stops[1].Offset, ApplyOpacity(gradient.Stops[1].Color, opacity) ?? gradient.Stops[1].Color));
    }

    private static bool TryCreateShadowShape(OfficeDrawingShape drawingShape, out OfficeDrawingShape shadowDrawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) {
            shadowDrawingShape = drawingShape;
            return false;
        }

        bool hasStroke = shape.Kind == OfficeShapeKind.Line ||
            (shape.StrokeColor.HasValue && shape.StrokeWidth > 0D && shape.StrokeColor.Value.A > 0);
        bool hasFill = shape.Kind != OfficeShapeKind.Line &&
            (shape.FillGradient != null || (shape.FillColor.HasValue && shape.FillColor.Value.A > 0));

        OfficeShape shadowShape = shape.Clone();
        shadowShape.Shadow = null;
        shadowShape.FillGradient = null;
        shadowShape.FillColor = hasFill || !hasStroke ? shadow.Color : null;
        shadowShape.FillOpacity = shadow.Opacity;
        shadowShape.StrokeColor = hasStroke ? shadow.Color : null;
        shadowShape.StrokeOpacity = shadow.Opacity;

        shadowDrawingShape = new OfficeDrawingShape(
            shadowShape,
            drawingShape.X + shadow.OffsetX,
            drawingShape.Y + shadow.OffsetY);
        return true;
    }
}
