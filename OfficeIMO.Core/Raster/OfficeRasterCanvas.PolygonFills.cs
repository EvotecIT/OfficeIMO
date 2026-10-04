using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    internal Func<double, double, OfficeColor> CreateContourPaint(IReadOnlyList<IReadOnlyList<OfficePoint>> contours,
        OfficeColor color, OfficeLinearGradient? linear, OfficeRadialGradient? radial) {
        if (linear == null && radial == null) return (_, _) => color;
        var points = new List<OfficePoint>();
        foreach (IReadOnlyList<OfficePoint> contour in contours) points.AddRange(contour);
        GetPolygonBounds(points, out double x, out double y, out double width, out double height);
        return (px, py) => {
            double nx = (px - x) / width, ny = (py - y) / height;
            if (radial != null) return InterpolateGradient(radial, ComputeRadialRatio(radial, nx, ny));
            double dx = linear!.EndX - linear.StartX, dy = linear.EndY - linear.StartY;
            double length = dx * dx + dy * dy;
            return length <= double.Epsilon ? linear.Stops[0].Color : InterpolateGradient(linear,
                Clamp(((nx - linear.StartX) * dx + (ny - linear.StartY) * dy) / length, 0D, 1D));
        };
    }

    private static IReadOnlyList<OfficePoint> RectanglePoints(double x, double y, double width, double height) =>
        new[] { new OfficePoint(x, y), new OfficePoint(x + width, y), new OfficePoint(x + width, y + height), new OfficePoint(x, y + height) };

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeColor color) =>
        FillContours(new[] { points }, color, OfficeFillRule.EvenOdd);

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeLinearGradient gradient) {
        GetPolygonBounds(points, out double x, out double y, out double width, out double height);
        double dx = gradient.EndX - gradient.StartX, dy = gradient.EndY - gradient.StartY;
        double lengthSquared = dx * dx + dy * dy;
        FillContourPaint(new[] { points }, OfficeFillRule.EvenOdd, (px, py) => {
            if (lengthSquared <= double.Epsilon) return gradient.Stops[0].Color;
            double ratio = ((((px - x) / width - gradient.StartX) * dx) + (((py - y) / height - gradient.StartY) * dy)) / lengthSquared;
            return InterpolateGradient(gradient, Clamp(ratio, 0D, 1D));
        });
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeRadialGradient gradient) {
        GetPolygonBounds(points, out double x, out double y, out double width, out double height);
        FillContourPaint(new[] { points }, OfficeFillRule.EvenOdd, (px, py) =>
            InterpolateGradient(gradient, ComputeRadialRatio(gradient, (px - x) / width, (py - y) / height)));
    }

    private static void GetPolygonBounds(IReadOnlyList<OfficePoint> points, out double x, out double y, out double width, out double height) {
        x = double.PositiveInfinity; y = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (OfficePoint point in points) { x = Math.Min(x, point.X); y = Math.Min(y, point.Y); right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y); }
        width = Math.Max(.0001D, right - x); height = Math.Max(.0001D, bottom - y);
    }
}
