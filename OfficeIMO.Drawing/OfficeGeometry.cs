using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable dependency-free geometry helpers shared by OfficeIMO renderers.
/// </summary>
public static class OfficeGeometry {
    private const double DefaultArrowheadWingAngleRadians = Math.PI / 7D;

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
    /// Calculates the triangular arrowhead points for a connector segment.
    /// </summary>
    /// <param name="tip">Arrowhead tip point.</param>
    /// <param name="from">A non-collapsed point behind the tip on the connector segment.</param>
    /// <param name="strokeWidth">Connector stroke width used to scale the arrowhead length.</param>
    /// <param name="points">Three arrowhead points: tip, first wing, and second wing.</param>
    /// <param name="minimumLength">Minimum arrowhead length in drawing units.</param>
    /// <param name="lengthMultiplier">Multiplier applied to <paramref name="strokeWidth"/>.</param>
    /// <param name="wingAngleRadians">Angle between the connector segment and each arrowhead wing.</param>
    /// <returns><see langword="true" /> when the segment and sizing inputs can produce an arrowhead.</returns>
    public static bool TryCreateArrowheadPoints(
        OfficePoint tip,
        OfficePoint from,
        double strokeWidth,
        out OfficePoint[] points,
        double minimumLength = 8D,
        double lengthMultiplier = 4D,
        double wingAngleRadians = DefaultArrowheadWingAngleRadians) {
        points = Array.Empty<OfficePoint>();
        if (!IsFinite(tip.X) ||
            !IsFinite(tip.Y) ||
            !IsFinite(from.X) ||
            !IsFinite(from.Y) ||
            !IsFinite(strokeWidth) ||
            !IsFinite(minimumLength) ||
            !IsFinite(lengthMultiplier) ||
            !IsFinite(wingAngleRadians)) {
            return false;
        }

        double dx = tip.X - from.X;
        double dy = tip.Y - from.Y;
        if ((dx * dx) + (dy * dy) <= 0D) {
            return false;
        }

        double length = Math.Max(strokeWidth * lengthMultiplier, minimumLength);
        if (length <= 0D) {
            return false;
        }

        double angle = Math.Atan2(dy, dx);
        points = new[] {
            tip,
            new OfficePoint(tip.X - Math.Cos(angle - wingAngleRadians) * length, tip.Y - Math.Sin(angle - wingAngleRadians) * length),
            new OfficePoint(tip.X - Math.Cos(angle + wingAngleRadians) * length, tip.Y - Math.Sin(angle + wingAngleRadians) * length)
        };
        return true;
    }

    /// <summary>
    /// Finds the terminal non-collapsed segment that should receive an arrowhead.
    /// </summary>
    /// <param name="points">Connector polyline points in drawing coordinates.</param>
    /// <param name="fromStart">When true, resolves the first segment; otherwise resolves the last segment.</param>
    /// <param name="tip">Resolved arrow tip point.</param>
    /// <param name="from">Resolved point behind the arrow tip.</param>
    /// <param name="tolerance">Minimum segment length considered non-collapsed.</param>
    /// <returns><see langword="true" /> when a non-collapsed terminal segment exists.</returns>
    public static bool TryGetArrowheadSegment(
        IReadOnlyList<OfficePoint> points,
        bool fromStart,
        out OfficePoint tip,
        out OfficePoint from,
        double tolerance = 1e-6D) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        if (points.Count < 2) {
            tip = default;
            from = default;
            return false;
        }

        double minimumDistance = NormalizeNonNegative(tolerance);
        if (fromStart) {
            tip = points[0];
            for (int i = 1; i < points.Count; i++) {
                if (Distance(tip, points[i]) > minimumDistance) {
                    from = points[i];
                    return true;
                }
            }
        } else {
            tip = points[points.Count - 1];
            for (int i = points.Count - 2; i >= 0; i--) {
                if (Distance(tip, points[i]) > minimumDistance) {
                    from = points[i];
                    return true;
                }
            }
        }

        from = default;
        return false;
    }

    /// <summary>
    /// Finds the terminal non-collapsed tuple segment that should receive an arrowhead.
    /// </summary>
    /// <param name="points">Connector polyline points in drawing coordinates.</param>
    /// <param name="fromStart">When true, resolves the first segment; otherwise resolves the last segment.</param>
    /// <param name="tip">Resolved arrow tip point.</param>
    /// <param name="from">Resolved point behind the arrow tip.</param>
    /// <param name="tolerance">Minimum segment length considered non-collapsed.</param>
    /// <returns><see langword="true" /> when a non-collapsed terminal segment exists.</returns>
    public static bool TryGetArrowheadSegment(
        IReadOnlyList<(double X, double Y)> points,
        bool fromStart,
        out (double X, double Y) tip,
        out (double X, double Y) from,
        double tolerance = 1e-6D) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        if (points.Count < 2) {
            tip = default;
            from = default;
            return false;
        }

        double minimumDistance = NormalizeNonNegative(tolerance);
        if (fromStart) {
            tip = points[0];
            for (int i = 1; i < points.Count; i++) {
                if (Distance(tip, points[i]) > minimumDistance) {
                    from = points[i];
                    return true;
                }
            }
        } else {
            tip = points[points.Count - 1];
            for (int i = points.Count - 2; i >= 0; i--) {
                if (Distance(tip, points[i]) > minimumDistance) {
                    from = points[i];
                    return true;
                }
            }
        }

        from = default;
        return false;
    }

    /// <summary>
    /// Resolves the point on a source rectangle boundary that faces a target rectangle.
    /// </summary>
    /// <param name="sourceLeft">Source rectangle left X coordinate.</param>
    /// <param name="sourceBottom">Source rectangle minimum Y coordinate.</param>
    /// <param name="sourceRight">Source rectangle right X coordinate.</param>
    /// <param name="sourceTop">Source rectangle maximum Y coordinate.</param>
    /// <param name="targetLeft">Target rectangle left X coordinate.</param>
    /// <param name="targetBottom">Target rectangle minimum Y coordinate.</param>
    /// <param name="targetRight">Target rectangle right X coordinate.</param>
    /// <param name="targetTop">Target rectangle maximum Y coordinate.</param>
    /// <returns>The source boundary point nearest the target rectangle by dominant center direction.</returns>
    public static OfficePoint ResolveRectangleBoundaryEndpoint(
        double sourceLeft,
        double sourceBottom,
        double sourceRight,
        double sourceTop,
        double targetLeft,
        double targetBottom,
        double targetRight,
        double targetTop) {
        double sourceMinX = Math.Min(sourceLeft, sourceRight);
        double sourceMaxX = Math.Max(sourceLeft, sourceRight);
        double sourceMinY = Math.Min(sourceBottom, sourceTop);
        double sourceMaxY = Math.Max(sourceBottom, sourceTop);
        double targetMinX = Math.Min(targetLeft, targetRight);
        double targetMaxX = Math.Max(targetLeft, targetRight);
        double targetMinY = Math.Min(targetBottom, targetTop);
        double targetMaxY = Math.Max(targetBottom, targetTop);

        double sourceCenterX = (sourceMinX + sourceMaxX) / 2D;
        double sourceCenterY = (sourceMinY + sourceMaxY) / 2D;
        double targetCenterX = (targetMinX + targetMaxX) / 2D;
        double targetCenterY = (targetMinY + targetMaxY) / 2D;
        double dx = targetCenterX - sourceCenterX;
        double dy = targetCenterY - sourceCenterY;

        return Math.Abs(dy) > Math.Abs(dx)
            ? new OfficePoint(sourceCenterX, dy >= 0D ? sourceMaxY : sourceMinY)
            : new OfficePoint(dx >= 0D ? sourceMaxX : sourceMinX, sourceCenterY);
    }

    /// <summary>
    /// Resolves coordinates on a source rectangle boundary that faces a target rectangle.
    /// </summary>
    /// <param name="sourceLeft">Source rectangle left X coordinate.</param>
    /// <param name="sourceBottom">Source rectangle minimum Y coordinate.</param>
    /// <param name="sourceRight">Source rectangle right X coordinate.</param>
    /// <param name="sourceTop">Source rectangle maximum Y coordinate.</param>
    /// <param name="targetLeft">Target rectangle left X coordinate.</param>
    /// <param name="targetBottom">Target rectangle minimum Y coordinate.</param>
    /// <param name="targetRight">Target rectangle right X coordinate.</param>
    /// <param name="targetTop">Target rectangle maximum Y coordinate.</param>
    /// <param name="x">Resolved source boundary X coordinate.</param>
    /// <param name="y">Resolved source boundary Y coordinate.</param>
    public static void ResolveRectangleBoundaryEndpoint(
        double sourceLeft,
        double sourceBottom,
        double sourceRight,
        double sourceTop,
        double targetLeft,
        double targetBottom,
        double targetRight,
        double targetTop,
        out double x,
        out double y) {
        OfficePoint point = ResolveRectangleBoundaryEndpoint(
            sourceLeft,
            sourceBottom,
            sourceRight,
            sourceTop,
            targetLeft,
            targetBottom,
            targetRight,
            targetTop);
        x = point.X;
        y = point.Y;
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

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && IsFinite(value) ? value : 0D;

    private static bool IsFinite(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value);
}
