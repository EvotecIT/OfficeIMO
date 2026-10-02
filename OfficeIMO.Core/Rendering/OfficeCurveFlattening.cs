using System;
using System.Collections.Generic;
namespace OfficeIMO.Drawing;

/// <summary>
/// Chooses how many straight segments a curve needs so its polygon stays within a fraction of
/// a device pixel of the true curve, subject to a bounded segment count, and builds common outlines.
/// </summary>
internal static class OfficeCurveFlattening {
    /// <summary>Target distance, in device pixels, between a flattened curve and the curve itself.</summary>
    internal const double Tolerance = 0.05;
    private const int MaximumSegments = 2048;

    /// <summary>Segments for a circular arc of <paramref name="radius"/> device pixels sweeping <paramref name="sweep"/> radians.</summary>
    internal static int ArcSegments(double radius, double sweep) {
        sweep = Math.Abs(sweep);
        if (!(radius > 0) || !(sweep > 0) || double.IsInfinity(radius) || double.IsInfinity(sweep)) return 1;
        var step = radius <= Tolerance ? Math.PI / 2 : 2 * Math.Acos(1 - Tolerance / radius);
        step = Math.Max(Math.PI / 720, Math.Min(Math.PI / 4, step));
        return (int)Math.Max(1D, Math.Min(MaximumSegments, Math.Ceiling(sweep / step - 0.000001)));
    }

    /// <summary>Segments for a quadratic Bézier whose control points are given in units of <paramref name="pixelsPerUnit"/>.</summary>
    internal static int QuadraticSegments(OfficePoint start, OfficePoint control, OfficePoint end, double pixelsPerUnit) {
        var deviation = Length(start.X - 2 * control.X + end.X, start.Y - 2 * control.Y + end.Y) * pixelsPerUnit;
        return SegmentsForDeviation(deviation / 4);
    }

    /// <summary>Segments for a cubic Bézier whose control points are given in units of <paramref name="pixelsPerUnit"/>.</summary>
    internal static int CubicSegments(OfficePoint start, OfficePoint control1, OfficePoint control2, OfficePoint end, double pixelsPerUnit) {
        var first = Length(start.X - 2 * control1.X + control2.X, start.Y - 2 * control1.Y + control2.Y);
        var second = Length(control1.X - 2 * control2.X + end.X, control1.Y - 2 * control2.Y + end.Y);
        return SegmentsForDeviation(Math.Max(first, second) * pixelsPerUnit * 0.75);
    }

    /// <summary>An ellipse outline as an open ring (the first point is not repeated).</summary>
    internal static List<OfficePoint> Ellipse(double cx, double cy, double rx, double ry, double pixelsPerUnit) {
        var segments = Math.Max(8, ArcSegments(Math.Max(rx, ry) * pixelsPerUnit, Math.PI * 2));
        var points = new List<OfficePoint>(segments);
        for (var i = 0; i < segments; i++) {
            var angle = Math.PI * 2 * i / segments;
            points.Add(new OfficePoint(cx + Math.Cos(angle) * rx, cy + Math.Sin(angle) * ry));
        }

        return points;
    }

    /// <summary>A circular arc as a polyline from <paramref name="startAngle"/> through <paramref name="sweep"/> radians.</summary>
    internal static List<OfficePoint> Arc(double cx, double cy, double radius, double startAngle, double sweep, double pixelsPerUnit) {
        var points = new List<OfficePoint>();
        AppendArc(points, cx, cy, radius, radius, startAngle, startAngle + sweep, pixelsPerUnit);
        return points;
    }

    /// <summary>A rounded rectangle outline as an open ring; zero radii give the four corners.</summary>
    internal static List<OfficePoint> RoundedRectangle(double x, double y, double width, double height, double rx, double ry, double pixelsPerUnit) {
        rx = Math.Max(0, Math.Min(rx, width / 2));
        ry = Math.Max(0, Math.Min(ry, height / 2));
        if (rx <= 0 || ry <= 0) {
            return new List<OfficePoint> { new OfficePoint(x, y), new OfficePoint(x + width, y), new OfficePoint(x + width, y + height), new OfficePoint(x, y + height) };
        }

        var points = new List<OfficePoint>();
        AppendArc(points, x + width - rx, y + ry, rx, ry, -Math.PI / 2, 0, pixelsPerUnit);
        AppendArc(points, x + width - rx, y + height - ry, rx, ry, 0, Math.PI / 2, pixelsPerUnit);
        AppendArc(points, x + rx, y + height - ry, rx, ry, Math.PI / 2, Math.PI, pixelsPerUnit);
        AppendArc(points, x + rx, y + ry, rx, ry, Math.PI, Math.PI * 1.5, pixelsPerUnit);
        return points;
    }

    private static void AppendArc(List<OfficePoint> points, double cx, double cy, double rx, double ry, double start, double end, double pixelsPerUnit) {
        var segments = ArcSegments(Math.Max(rx, ry) * pixelsPerUnit, end - start);
        for (var i = 0; i <= segments; i++) {
            var angle = start + (end - start) * i / segments;
            points.Add(new OfficePoint(cx + Math.Cos(angle) * rx, cy + Math.Sin(angle) * ry));
        }
    }

    private static int SegmentsForDeviation(double deviation) {
        if (!(deviation > 0)) return 1;
        return (int)Math.Max(1D, Math.Min(MaximumSegments, Math.Ceiling(Math.Sqrt(deviation / Tolerance))));
    }

    private static double Length(double x, double y) => Math.Sqrt(x * x + y * y);
}
