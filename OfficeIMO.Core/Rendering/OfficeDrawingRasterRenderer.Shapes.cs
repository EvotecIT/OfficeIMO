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
                if (strokeWidth > 0D) RenderTransformedLine(canvas, drawingShape, scale, stroke ?? fill ?? OfficeColor.Black, strokeGradient, strokeRadialGradient, strokeWidth);
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
            OfficePoint a = TransformShapePoint(drawingShape, shape.Points[0], scale);
            OfficePoint b = TransformShapePoint(drawingShape, shape.Points[1], scale);
            OfficeColor startColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, shape.Points[0], shape.Points[1], shape.Points[0]);
            OfficeColor endColor = SampleLineMarkerColor(color, strokeGradient, strokeRadialGradient, shape.Points[0], shape.Points[1], shape.Points[1]);
            StrokeTransformedPathContours(canvas, drawingShape, new[] { new OfficeFlattenedPathContour(shape.Points, false) }, scale, color, strokeGradient, strokeRadialGradient);
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
        if (fillRadialGradient != null) {
            fillRadialGradient = TransformShapeFillGradient(drawingShape, scale,
                new[] { (IReadOnlyList<OfficePoint>)points }, fillRadialGradient);
            canvas.FillRadialGradientPolygon(points, fillRadialGradient);
        }
        else if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
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
                if (fillGradient != null) {
                    fillGradient = TransformShapeFillGradient(drawingShape, scale,
                        closedContours, fillGradient);
                }
                if (fillRadialGradient != null) {
                    fillRadialGradient = TransformShapeFillGradient(drawingShape, scale, closedContours, fillRadialGradient);
                }
                if (fillRadialGradient != null || fillGradient != null) {
                    FillGradientPathContours(canvas, closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            StrokeTransformedPathContours(canvas, drawingShape, contours, scale, stroke, strokeGradient, strokeRadialGradient);
            RenderPathMarkers(canvas, shape, contours, stroke ?? GetGradientFallbackStroke(strokeGradient, strokeRadialGradient) ?? OfficeColor.Black, scale, strokeGradient, strokeRadialGradient, 0D, 0D, shape.Width, shape.Height, point => TransformShapePoint(drawingShape, point, scale));
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
            DrawGradientOrSolidPolyline(canvas, new[] { start, end }, color, strokeGradient, strokeRadialGradient, strokeWidth, shape, close: false, shape.StrokeLineCap);
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

    private static OfficeColor SampleLineMarkerColor(OfficeColor fallback, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, OfficePoint start, OfficePoint end, OfficePoint samplePoint) {
        double left = Math.Min(start.X, end.X);
        double top = Math.Min(start.Y, end.Y);
        double width = Math.Abs(end.X - start.X);
        double height = Math.Abs(end.Y - start.Y);
        return SampleStrokeGradient(strokeGradient, strokeRadialGradient, left, top, width, height, samplePoint.X, samplePoint.Y) ?? fallback;
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
        double.IsNaN(ratio) ? OfficeColor.Transparent : InterpolateGradientStops(gradient.Stops, ratio);

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

    private static double ComputeRadialRatio(OfficeRadialGradient gradient, double x, double y) => gradient.SampleRatio(x, y);

    private static double Clamp(double value, double min, double max) =>
        value < min ? min : value > max ? max : value;

    private static OfficeLineMarker? ScaleLineMarker(OfficeLineMarker? marker, double scale) =>
        marker == null ? null : new OfficeLineMarker(marker.Kind, marker.Width * scale, marker.Length * scale);

    private static void RenderPolygon(OfficeRasterCanvas canvas, OfficeShape shape, double x, double y, double scale, OfficeColor? fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? stroke, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth) {
        List<OfficePoint> points = OffsetPoints(shape.Points, x, y, scale);
        if (fillRadialGradient != null) canvas.FillRadialGradientPolygon(points, fillRadialGradient);
        else if (fillGradient != null) canvas.FillLinearGradientPolygon(points, fillGradient);
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
                    var coordinates = NormalizePaintCoordinates(new OfficeTransform(
                        shape.Width * scale, 0D, 0D, shape.Height * scale, x, y), closedContours);
                    if (fillGradient != null) fillGradient = fillGradient.TransformCoordinates(coordinates);
                    if (fillRadialGradient != null) fillRadialGradient = fillRadialGradient.TransformCoordinates(coordinates);
                    FillGradientPathContours(canvas, closedContours, fillGradient, fillRadialGradient, shape.FillRule);
                } else {
                    FillPathContours(canvas, closedContours, fill!.Value, shape.FillRule);
                }
            }
        }

        if ((stroke.HasValue || strokeGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
            StrokePathContours(canvas, contours, stroke, strokeGradient, strokeRadialGradient, strokeWidth, shape,
                paintBounds: (x, y, shape.Width * scale, shape.Height * scale));
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

}
