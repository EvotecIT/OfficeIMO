using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

// Rectangles and convex paths use half-planes; other paths retain their fill geometry.
internal readonly partial struct OfficeTextInkClip {
    private readonly IReadOnlyList<OfficePoint>? _polygon;
    private readonly double _orientation;
    internal IReadOnlyList<IReadOnlyList<OfficePoint>>? FilledContours { get; }
    internal OfficeFillRule FillRule { get; }
    internal OfficeTextInkClip(double left, double top, double width, double height,
        bool horizontal, bool vertical, OfficeTransform inverse) {
        this = default; _polygon = null; _orientation = 0D;
        Left = left; Top = top; Right = left + width; Bottom = top + height;
        Horizontal = horizontal; Vertical = vertical; Inverse = inverse;
    }
    private double Left { get; }
    private double Top { get; }
    private double Right { get; }
    private double Bottom { get; }
    private bool Horizontal { get; }
    private bool Vertical { get; }
    private OfficeTransform Inverse { get; }

    // Sutherland-Hodgman intersection with the enabled half-planes. Individual
    // contours retain their orientation, but their bounds do not resolve winding
    // cancellation between overlapping contours or establish pixel coverage.
    internal List<OfficePoint> Apply(List<OfficePoint> points, ref bool clipped, ref long remainingWork,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (FilledContours != null) throw new InvalidOperationException("Filled path clips must be intersected during filled-ink measurement.");
        if (_polygon == null && ((Horizontal && Right <= Left) || (Vertical && Bottom <= Top))) {
            clipped |= points.Count > 0;
            return new List<OfficePoint>();
        }
        for (int edge = 0; edge < (_polygon?.Count ?? 4) && points.Count > 0; edge++) {
            if (_polygon == null && ((edge < 2 && !Horizontal) || (edge >= 2 && !Vertical))) continue;
            if (points.Count > remainingWork) throw new NotSupportedException("Text ink clipping exceeds its point-work limit.");
            remainingWork -= points.Count;
            var output = new List<OfficePoint>();
            OfficePoint previous = points[points.Count - 1];
            double previousDistance = Distance(previous, edge);
            for (int i = 0; i < points.Count; i++) {
                if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                OfficePoint current = points[i];
                double distance = Distance(current, edge);
                bool inside = distance >= 0D, previousInside = previousDistance >= 0D;
                if (inside != previousInside) {
                    double ratio = previousDistance == 0D ? 0D : 1D / (1D + Math.Abs(distance / previousDistance));
                    output.Add(new OfficePoint(previous.X * (1D - ratio) + current.X * ratio,
                        previous.Y * (1D - ratio) + current.Y * ratio));
                }
                if (inside) output.Add(current); else clipped = true;
                previous = current; previousDistance = distance;
            }
            points = output;
        }
        return points;
    }

    private double Distance(OfficePoint point, int edge) {
        double distance;
        if (_polygon != null) {
            OfficePoint start = _polygon[edge], end = _polygon[(edge + 1) % _polygon.Count];
            distance = Cross(start, end, point) * _orientation;
        } else {
            OfficePoint local = Inverse.TransformPoint(point);
            distance = edge switch { 0 => local.X - Left, 1 => Right - local.X, 2 => local.Y - Top, _ => Bottom - local.Y };
        }
        if (double.IsNaN(distance) || double.IsInfinity(distance))
            throw new NotSupportedException("Text ink clipping requires finite geometry.");
        return distance;
    }
}
