using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable dependency-free geometry helpers shared by OfficeIMO renderers.
/// </summary>
public static class OfficeGeometry {
    /// <summary>
    /// Calculates the Euclidean distance between two points.
    /// </summary>
    public static double Distance(double x1, double y1, double x2, double y2) {
        double dx = x2 - x1;
        double dy = y2 - y1;
        return Math.Sqrt((dx * dx) + (dy * dy));
    }

    /// <summary>
    /// Calculates the Euclidean distance between two points.
    /// </summary>
    public static double Distance(OfficePoint from, OfficePoint to) =>
        Distance(from.X, from.Y, to.X, to.Y);

    /// <summary>
    /// Calculates the Euclidean distance between two tuple points.
    /// </summary>
    public static double Distance((double X, double Y) from, (double X, double Y) to) =>
        Distance(from.X, from.Y, to.X, to.Y);

    /// <summary>
    /// Returns the point at the supplied normalized position along a polyline.
    /// </summary>
    /// <param name="points">Polyline points in drawing coordinates.</param>
    /// <param name="position">Normalized position where 0 is the first point and 1 is the last point.</param>
    public static OfficePoint InterpolatePolyline(IReadOnlyList<OfficePoint> points, double position) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        if (points.Count == 0) {
            return default;
        }

        if (points.Count == 1) {
            return points[0];
        }

        double total = 0D;
        for (int i = 1; i < points.Count; i++) {
            total += Distance(points[i - 1], points[i]);
        }

        if (total <= 0D) {
            return points[0];
        }

        double target = total * ClampPosition(position);
        double traversed = 0D;
        for (int i = 1; i < points.Count; i++) {
            OfficePoint from = points[i - 1];
            OfficePoint to = points[i];
            double segment = Distance(from, to);
            if (segment <= 0D) {
                continue;
            }

            if (traversed + segment >= target) {
                double t = (target - traversed) / segment;
                return new OfficePoint(
                    from.X + ((to.X - from.X) * t),
                    from.Y + ((to.Y - from.Y) * t));
            }

            traversed += segment;
        }

        return points[points.Count - 1];
    }

    /// <summary>
    /// Returns the tuple point at the supplied normalized position along a polyline.
    /// </summary>
    /// <param name="points">Polyline points in drawing coordinates.</param>
    /// <param name="position">Normalized position where 0 is the first point and 1 is the last point.</param>
    public static (double X, double Y) InterpolatePolyline(IReadOnlyList<(double X, double Y)> points, double position) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        if (points.Count == 0) {
            return default;
        }

        if (points.Count == 1) {
            return points[0];
        }

        double total = 0D;
        for (int i = 1; i < points.Count; i++) {
            total += Distance(points[i - 1], points[i]);
        }

        if (total <= 0D) {
            return points[0];
        }

        double target = total * ClampPosition(position);
        double traversed = 0D;
        for (int i = 1; i < points.Count; i++) {
            (double X, double Y) from = points[i - 1];
            (double X, double Y) to = points[i];
            double segment = Distance(from, to);
            if (segment <= 0D) {
                continue;
            }

            if (traversed + segment >= target) {
                double t = (target - traversed) / segment;
                return (
                    from.X + ((to.X - from.X) * t),
                    from.Y + ((to.Y - from.Y) * t));
            }

            traversed += segment;
        }

        return points[points.Count - 1];
    }

    private static double ClampPosition(double position) {
        if (double.IsNaN(position)) {
            return 0D;
        }

        if (position < 0D) {
            return 0D;
        }

        return position > 1D ? 1D : position;
    }
}
