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
        if (stroke.HasValue) canvas.DrawPolygon(points, stroke.Value, strokeWidth);
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

    private static OfficeColor? ApplyOpacity(OfficeColor? color, double? opacity) {
        if (!color.HasValue) return null;
        if (!opacity.HasValue) return color;
        double clamped = opacity.Value < 0D ? 0D : opacity.Value > 1D ? 1D : opacity.Value;
        return OfficeColor.FromRgba(color.Value.R, color.Value.G, color.Value.B, (byte)Math.Round(color.Value.A * clamped));
    }
}
