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
    /// Calculates the perpendicular half-offset for two parallel lines separated by the supplied distance.
    /// </summary>
    /// <param name="x1">Source line start X coordinate.</param>
    /// <param name="y1">Source line start Y coordinate.</param>
    /// <param name="x2">Source line end X coordinate.</param>
    /// <param name="y2">Source line end Y coordinate.</param>
    /// <param name="separation">Distance between the two parallel line centers.</param>
    /// <param name="offsetX">Perpendicular X offset from the source line center to either parallel line.</param>
    /// <param name="offsetY">Perpendicular Y offset from the source line center to either parallel line.</param>
    /// <returns><see langword="true" /> when the source line has usable length and finite coordinates.</returns>
    public static bool TryGetParallelLineOffsets(
        double x1,
        double y1,
        double x2,
        double y2,
        double separation,
        out double offsetX,
        out double offsetY) {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt((dx * dx) + (dy * dy));
        if (!IsFinite(length) || length <= 0D || !IsFinite(separation)) {
            offsetX = 0D;
            offsetY = 0D;
            return false;
        }

        double half = separation / 2D;
        offsetX = -dy / length * half;
        offsetY = dx / length * half;
        return true;
    }

    /// <summary>
    /// Converts degrees to radians.
    /// </summary>
    /// <param name="degrees">Angle in degrees.</param>
    /// <returns>The same angle in radians.</returns>
    public static double DegreesToRadians(double degrees) => degrees * Math.PI / 180D;

    /// <summary>
    /// Converts radians to degrees.
    /// </summary>
    /// <param name="radians">Angle in radians.</param>
    /// <returns>The same angle in degrees.</returns>
    public static double RadiansToDegrees(double radians) => radians * 180D / Math.PI;

    /// <summary>
    /// Rotates a point around the supplied center using radians in raster coordinate space.
    /// </summary>
    /// <param name="point">Point to rotate.</param>
    /// <param name="centerX">Rotation center X coordinate.</param>
    /// <param name="centerY">Rotation center Y coordinate.</param>
    /// <param name="rotationRadians">Rotation angle in radians.</param>
    /// <returns>The rotated point.</returns>
    public static OfficePoint RotatePoint(OfficePoint point, double centerX, double centerY, double rotationRadians) {
        if (Math.Abs(rotationRadians) < 0.000001D) {
            return point;
        }

        double cos = Math.Cos(rotationRadians);
        double sin = Math.Sin(rotationRadians);
        double dx = point.X - centerX;
        double dy = point.Y - centerY;
        return new OfficePoint(
            centerX + (dx * cos) - (dy * sin),
            centerY + (dx * sin) + (dy * cos));
    }

    /// <summary>
    /// Rotates a tuple point around the supplied center using radians in raster coordinate space.
    /// </summary>
    /// <param name="point">Point to rotate.</param>
    /// <param name="centerX">Rotation center X coordinate.</param>
    /// <param name="centerY">Rotation center Y coordinate.</param>
    /// <param name="rotationRadians">Rotation angle in radians.</param>
    /// <returns>The rotated tuple point.</returns>
    public static (double X, double Y) RotatePoint((double X, double Y) point, double centerX, double centerY, double rotationRadians) {
        OfficePoint rotated = RotatePoint(new OfficePoint(point.X, point.Y), centerX, centerY, rotationRadians);
        return (rotated.X, rotated.Y);
    }

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

    private static bool IsFinite(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value);
}
