using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderTransformedShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        bool hasFillArea = HasTransformedFillArea(shape);
        OfficeColor? fill = hasFillArea ? ApplyOpacity(shape.FillColor, shape.FillOpacity) : null;
        OfficeLinearGradient? fillGradient = !hasFillArea || shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity);
        OfficeRadialGradient? fillRadialGradient = !hasFillArea || shape.FillRadialGradient == null ? null : ApplyOpacity(shape.FillRadialGradient, shape.FillOpacity);

        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        OfficeLinearGradient? strokeGradient = shape.StrokeGradient == null ? null : ApplyOpacity(shape.StrokeGradient, shape.StrokeOpacity);
        OfficeRadialGradient? strokeRadialGradient = shape.StrokeRadialGradient == null ? null : ApplyOpacity(shape.StrokeRadialGradient, shape.StrokeOpacity);
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.RoundedRectangle:
            case OfficeShapeKind.Ellipse:
                RenderTransformedClosedContour(canvas, drawingShape, scale, CreateShapeContour(shape, GetShapePixelScale(drawingShape, scale)), fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D && (stroke.HasValue || strokeGradient != null || strokeRadialGradient != null))
                    RenderTransformedLine(canvas, drawingShape, scale, stroke ?? OfficeColor.Transparent, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderTransformedClosedContour(canvas, drawingShape, scale, shape.Points, fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderTransformedPath(canvas, drawingShape, scale, fill, fillGradient, fillRadialGradient, stroke, strokeGradient, strokeRadialGradient, strokeWidth);
                break;
        }
    }

    private static bool HasTransformedFillArea(OfficeShape shape) {
        if (shape.Width == 0D || shape.Height == 0D) return false;
        OfficeTransform transform = shape.Transform ?? OfficeTransform.Identity;
        double scale = Math.Max(Math.Max(Math.Abs(transform.M11), Math.Abs(transform.M12)),
            Math.Max(Math.Abs(transform.M21), Math.Abs(transform.M22)));
        if (scale == 0D) return false;
        // A singular transform can leave a diagonal bounding box with positive
        // width and height, but its fill still has no area. Keep stroke handling
        // independent and do not attempt to invert that collapsed color field.
        return (transform.M11 / scale) * (transform.M22 / scale)
            - (transform.M12 / scale) * (transform.M21 / scale) != 0D;
    }

    private static void RenderTransformedLine(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor color, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count >= 2) {
            StrokeTransformedPathContours(canvas, drawingShape, new[] { new OfficeFlattenedPathContour(shape.Points, false) }, scale, color, strokeGradient, strokeRadialGradient);
            RenderLineMarkers(canvas, shape, shape.Points[0], shape.Points[1], color, strokeGradient, strokeRadialGradient, GetRasterTransform(drawingShape, scale));
        }
    }

    private static void RenderTransformedClosedContour(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, IReadOnlyList<OfficePoint> contour, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        if (contour.Count < 3) {
            return;
        }

        List<OfficePoint> points = TransformShapePoints(drawingShape, contour, scale);
        if (fillRadialGradient != null || fillGradient != null)
            FillShapeGradientContours(canvas, drawingShape.Shape, drawingShape.X, drawingShape.Y, scale,
                new[] { (IReadOnlyList<OfficePoint>)points }, fillGradient, fillRadialGradient, drawingShape.Shape.FillRule);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        StrokeTransformedPathContours(canvas, drawingShape, new[] { new OfficeFlattenedPathContour(contour, true) }, scale, stroke, strokeGradient, strokeRadialGradient);
    }

    private static void RenderTransformedPath(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        OfficeShape shape = drawingShape.Shape;
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D, pixelsPerUnit: GetShapePixelScale(drawingShape, scale));
        if (fillRadialGradient != null || fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                // Filled contours close implicitly; retain their open state for stroking.
                if (contours[i].Points.Count >= 3) {
                    closedContours.Add(TransformShapePoints(drawingShape, contours[i].Points, scale));
                }
            }

            if (closedContours.Count > 0) {
                if (fillRadialGradient != null || fillGradient != null) {
                    FillShapeGradientContours(canvas, shape, drawingShape.X, drawingShape.Y, scale,
                        closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            StrokeTransformedPathContours(canvas, drawingShape, contours, scale, stroke, strokeGradient, strokeRadialGradient);
            RenderPathMarkers(canvas, shape, contours, stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient) ?? OfficeColor.Black,
                strokeGradient, strokeRadialGradient, (0D, 0D, shape.Width, shape.Height), OfficeTransform.Identity, GetRasterTransform(drawingShape, scale));
        }
    }

    private static void RenderLine(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor color, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        if (shape.Points.Count >= 2) {
            OfficePoint a = shape.Points[0];
            OfficePoint b = shape.Points[1];
            OfficePoint start = new OfficePoint(x + (a.X * scale), y + (a.Y * scale));
            OfficePoint end = new OfficePoint(x + (b.X * scale), y + (b.Y * scale));
            DrawGradientOrSolidPolyline(canvas, new[] { start, end }, color, strokeGradient, strokeRadialGradient, strokeWidth, shape, close: false, shape.StrokeLineCap);
            RenderLineMarkers(canvas, shape, a, b, color, strokeGradient, strokeRadialGradient, OfficeTransform.Scale(scale, scale).Then(OfficeTransform.Translate(x, y)));
        }
    }

    private static void DrawGradientOrSolidPolyline(
        OfficeRasterCanvas canvas,
        IReadOnlyList<OfficePoint> points,
        OfficeColor? stroke,
        OfficeLinearGradient? strokeGradient,
        OfficeRadialGradient? strokeRadialGradient,
        double strokeWidth,
        OfficeShape shape,
        bool close,
        OfficeStrokeLineCap? lineCap = null) {
        StrokePathContours(canvas, new[] { new OfficeFlattenedPathContour(points, close) }, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape, lineCap);
    }

    private static void StrokePathContours(OfficeRasterCanvas canvas, IReadOnlyList<OfficeFlattenedPathContour> contours,
        OfficeColor? stroke, OfficeLinearGradient? linear, OfficeRadialGradient? radial, double width, OfficeShape shape, OfficeStrokeLineCap? cap = null,
        (double X, double Y, double Width, double Height)? paintBounds = null) {
        if (width <= 0D || (stroke == null && linear == null && radial == null)) return;
        var allPoints = new List<OfficePoint>();
        foreach (OfficeFlattenedPathContour contour in contours) allPoints.AddRange(contour.Points);
        if (allPoints.Count == 0) return;
        GetPointBounds(allPoints, out double x, out double y, out double w, out double h);
        if (paintBounds.HasValue) (x, y, w, h) = paintBounds.Value;
        double strokeScale = shape.StrokeWidth > 0D ? width / shape.StrokeWidth : 1D;
        IReadOnlyList<double>? pattern = shape.StrokeDashStyle.GetDashPattern(width);
        if (shape.StrokeDashArray.Count > 0) {
            var exact = new double[shape.StrokeDashArray.Count];
            for (int i = 0; i < exact.Length; i++) exact[i] = shape.StrokeDashArray[i] * strokeScale;
            pattern = exact;
        }
        canvas.StrokeContours(contours, width, cap ?? shape.StrokeLineCap ?? OfficeStrokeLineCap.Round,
            shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Round, shape.StrokeMiterLimit, pattern, shape.StrokeDashOffset * strokeScale,
            (px, py) => SampleStrokeGradient(linear, radial, x, y, w, h, px, py) ?? stroke ?? OfficeColor.Transparent);
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

    private static OfficeColor? SampleStrokeGradient(OfficeLinearGradient? linearGradient, OfficeRadialGradient? radialGradient, double x, double y, double width, double height, double sampleX, double sampleY) {
        width = Math.Max(width, 0.0001D);
        height = Math.Max(height, 0.0001D);
        double nx = (sampleX - x) / width;
        double ny = (sampleY - y) / height;
        if (radialGradient != null) {
            return InterpolateGradient(radialGradient, OfficeRasterCanvas.ComputeRadialRatio(radialGradient, nx, ny));
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
        InterpolateGradientStops(gradient.Stops, ratio, gradient.ColorInterpolation);

    private static OfficeColor InterpolateGradient(OfficeRadialGradient gradient, double ratio) =>
        double.IsNaN(ratio) ? gradient.OutsideColor ?? OfficeColor.Transparent : InterpolateGradientStops(gradient.Stops, ratio, gradient.ColorInterpolation);

    private static OfficeColor InterpolateGradientStops(IReadOnlyList<OfficeGradientStop> stops, double ratio, OfficeGradientColorInterpolation interpolation) =>
        OfficeRasterCanvas.InterpolateGradientStops(stops, ratio, separateAlpha: true, interpolation: interpolation);

    private static double Clamp(double value, double min, double max) =>
        value < min ? min : value > max ? max : value;

    private static void RenderPolygon(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        List<OfficePoint> points = OffsetPoints(shape.Points, x, y, scale);
        if (fillRadialGradient != null || fillGradient != null)
            FillShapeGradientContours(canvas, shape, x / scale, y / scale, scale,
                new[] { (IReadOnlyList<OfficePoint>)points }, fillGradient, fillRadialGradient, shape.FillRule);
        else if (fill.HasValue) canvas.FillPolygon(points, fill.Value);
        DrawGradientOrSolidPolyline(canvas, points, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape, close: true);
    }

    private static void RenderPath(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, x, y, scale);
        if (fillRadialGradient != null || fillGradient != null || fill.HasValue) {
            List<IReadOnlyList<OfficePoint>> closedContours = new List<IReadOnlyList<OfficePoint>>();
            for (int i = 0; i < contours.Count; i++) {
                // SVG and PDF fills close open contours without closing their strokes.
                if (contours[i].Points.Count >= 3) {
                    closedContours.Add(contours[i].Points);
                }
            }

            if (closedContours.Count > 0) {
                if (fillRadialGradient != null || fillGradient != null) {
                    FillShapeGradientContours(canvas, shape, x / scale, y / scale, scale,
                        closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            StrokePathContours(canvas, contours, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape,
                paintBounds: (x, y, shape.Width * scale, shape.Height * scale));
            RenderPathMarkers(canvas, shape, contours, stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient) ?? OfficeColor.Black,
                strokeGradient, strokeRadialGradient, (x, y, shape.Width * scale, shape.Height * scale),
                OfficeTransform.Translate(-x, -y).Then(OfficeTransform.Scale(1D / scale, 1D / scale)),
                OfficeTransform.Scale(scale, scale).Then(OfficeTransform.Translate(x, y)));
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

}
