using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

// A rectangular scene clip expressed in its own space. Inverse maps measured
// outline points from the inspection space to that clip's local coordinates.
internal readonly struct OfficeTextInkClip {
    internal OfficeTextInkClip(double left, double top, double width, double height,
        bool horizontal, bool vertical, OfficeTransform inverse) {
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
        if ((Horizontal && Right <= Left) || (Vertical && Bottom <= Top)) {
            clipped |= points.Count > 0;
            return new List<OfficePoint>();
        }
        for (int edge = 0; edge < 4 && points.Count > 0; edge++) {
            if ((edge < 2 && !Horizontal) || (edge >= 2 && !Vertical)) continue;
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
        OfficePoint local = Inverse.TransformPoint(point);
        double distance = edge switch { 0 => local.X - Left, 1 => Right - local.X, 2 => local.Y - Top, _ => Bottom - local.Y };
        if (double.IsNaN(distance) || double.IsInfinity(distance))
            throw new NotSupportedException("Text ink clipping requires finite geometry.");
        return distance;
    }
}
