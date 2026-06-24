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
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape shape) {
                RenderShape(canvas, shape, scale);
            } else if (element is OfficeDrawingText text) {
                RenderText(canvas, text, scale);
            }
        }

        return image;
    }

    /// <summary>
    /// Renders a drawing to PNG bytes.
    /// </summary>
    public static byte[] ToPng(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) =>
        OfficePngWriter.Encode(Render(drawing, scale, background));

    private static void RenderShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
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
        double strokeWidth = Math.Max(1D, shape.StrokeWidth * scale);

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.RoundedRectangle:
                if (shape.FillGradient != null) {
                    canvas.FillLinearGradientRectangle(x, y, width, height, shape.FillGradient);
                } else if (fill.HasValue) {
                    canvas.FillRectangle(x, y, width, height, fill.Value);
                }

                if (stroke.HasValue) canvas.DrawRectangle(x, y, width, height, stroke.Value, strokeWidth);
                break;
            case OfficeShapeKind.Ellipse:
                if (fill.HasValue) canvas.FillEllipse(x, y, width, height, fill.Value);
                if (stroke.HasValue) canvas.DrawEllipse(x, y, width, height, stroke.Value, strokeWidth);
                break;
            case OfficeShapeKind.Line:
                RenderLine(canvas, shape, x, y, scale, stroke ?? fill ?? OfficeColor.Black, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderPolygon(canvas, shape, x, y, scale, fill, stroke, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderPath(canvas, shape, x, y, scale, fill, stroke, strokeWidth, shape.StrokeDashStyle);
                break;
        }
    }

    private static void RenderText(OfficeRasterCanvas canvas, OfficeDrawingText text, double scale) {
        if (!text.WrapText && !text.ShrinkToFit && Math.Abs(text.RotationDegrees) <= 0.000001D && text.VerticalAlignment == OfficeTextVerticalAlignment.Top) {
            canvas.DrawText(
                text.Text,
                text.X * scale,
                text.Y * scale,
                text.Width * scale,
                text.Height * scale,
                text.Color ?? OfficeColor.Black,
                text.Font.Size * scale,
                text.Alignment,
                text.Font.Style,
                text.Font.FamilyName);
            return;
        }

        double fontSize = Math.Max(1D, text.Font.Size * scale);
        double textWidth = text.Width * scale;
        double textHeight = text.Height * scale;
        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, (text.LineHeight.Value * scale) / fontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D, fontSize);
        Func<string?, double, double> measure = (value, size) => canvas.MeasureText(value, size, text.Font.FamilyName);
        OfficeTextBlockLayout layout = text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedText(
                text.Text,
                fontSize,
                textWidth,
                textHeight,
                lineHeightFactor,
                minimumFontSize,
                measure)
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.Text,
                fontSize,
                textWidth,
                textHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                wrap: text.WrapText,
                shrinkToFit: text.ShrinkToFit);
        OfficeTextBlockRenderer.DrawRasterTextBlock(
            canvas,
            layout,
            text.X * scale,
            text.Y * scale,
            textWidth,
            textHeight,
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
            fontFamily: text.Font.FamilyName);
    }

    private static void RenderTransformedShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        OfficeColor? fill = ApplyOpacity(shape.FillColor, shape.FillOpacity);
        if (!fill.HasValue && shape.FillGradient != null && shape.FillGradient.Stops.Count > 0) {
            fill = ApplyOpacity(shape.FillGradient.Stops[0].Color, shape.FillOpacity);
        }

        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        double strokeWidth = Math.Max(1D, shape.StrokeWidth * scale);

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.RoundedRectangle:
            case OfficeShapeKind.Ellipse:
                RenderTransformedClosedContour(canvas, drawingShape, scale, CreateShapeContour(shape), fill, stroke, strokeWidth);
                break;
            case OfficeShapeKind.Line:
                RenderTransformedLine(canvas, drawingShape, scale, stroke ?? fill ?? OfficeColor.Black, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderTransformedClosedContour(canvas, drawingShape, scale, shape.Points, fill, stroke, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderTransformedPath(canvas, drawingShape, scale, fill, stroke, strokeWidth, shape.StrokeDashStyle);
                break;
        }
    }

    private static void RenderTransformedLine(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor color, double strokeWidth) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count >= 2) {
            OfficePoint a = TransformShapePoint(drawingShape, shape.Points[0], scale);
            OfficePoint b = TransformShapePoint(drawingShape, shape.Points[1], scale);
            canvas.DrawStyledLine(a.X, a.Y, b.X, b.Y, color, strokeWidth, shape.StrokeDashStyle);
        }
    }

    private static void RenderTransformedClosedContour(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, IReadOnlyList<OfficePoint> contour, OfficeColor? fill, OfficeColor? stroke, double strokeWidth) {
        if (contour.Count < 3) {
            return;
        }

        List<OfficePoint> points = TransformShapePoints(drawingShape, contour, scale);
        if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        if (stroke.HasValue) canvas.DrawStyledPolygon(points, stroke.Value, strokeWidth, drawingShape.Shape.StrokeDashStyle);
    }

    private static void RenderTransformedPath(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        OfficeShape shape = drawingShape.Shape;
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D);
        if (fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(TransformShapePoints(drawingShape, contours[i].Points, scale));
                }
            }

            if (closedContours.Count > 0) {
                canvas.FillPolygonsEvenOdd(closedContours, fill.Value);
            }
        }

        if (stroke.HasValue) {
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
            canvas.DrawStyledLine(x + (a.X * scale), y + (a.Y * scale), x + (b.X * scale), y + (b.Y * scale), color, strokeWidth, shape.StrokeDashStyle);
        }
    }

    private static void RenderPolygon(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeColor? stroke, double strokeWidth) {
        List<OfficePoint> points = OffsetPoints(shape.Points, x, y, scale);
        if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        if (stroke.HasValue) canvas.DrawStyledPolygon(points, stroke.Value, strokeWidth, shape.StrokeDashStyle);
    }

    private static void RenderPath(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, x, y, scale);
        if (fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                if (contours[i].Closed && contours[i].Points.Count >= 3) {
                    closedContours.Add(contours[i].Points);
                }
            }

            if (closedContours.Count > 0) {
                canvas.FillPolygonsEvenOdd(closedContours, fill.Value);
            }
        }

        if (stroke.HasValue) {
            for (int i = 0; i < contours.Count; i++) {
                IReadOnlyList<OfficePoint> points = contours[i].Closed
                    ? CloseContour(contours[i].Points)
                    : contours[i].Points;
                canvas.DrawStyledPolyline(points, stroke.Value, strokeWidth, dashStyle);
            }
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
}
