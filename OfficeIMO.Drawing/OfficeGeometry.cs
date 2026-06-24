using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable dependency-free geometry helpers shared by OfficeIMO renderers.
/// </summary>
public static class OfficeGeometry {
    private const double DefaultArrowheadWingAngleRadians = Math.PI / 7D;
    private const double GeometryTolerance = 1e-9D;

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
    /// Determines whether two line segments intersect, including boundary touches and collinear overlap.
    /// </summary>
    public static bool SegmentsIntersect(
        (double X, double Y) firstStart,
        (double X, double Y) firstEnd,
        (double X, double Y) secondStart,
        (double X, double Y) secondEnd) {
        double o1 = Orientation(firstStart, firstEnd, secondStart);
        double o2 = Orientation(firstStart, firstEnd, secondEnd);
        double o3 = Orientation(secondStart, secondEnd, firstStart);
        double o4 = Orientation(secondStart, secondEnd, firstEnd);

        if (o1 * o2 < 0D && o3 * o4 < 0D) {
            return true;
        }

        return IsZero(o1) && OnSegment(firstStart, secondStart, firstEnd) ||
               IsZero(o2) && OnSegment(firstStart, secondEnd, firstEnd) ||
               IsZero(o3) && OnSegment(secondStart, firstStart, secondEnd) ||
               IsZero(o4) && OnSegment(secondStart, firstEnd, secondEnd);
    }

    /// <summary>
    /// Determines whether two line segments intersect, including boundary touches and collinear overlap.
    /// </summary>
    public static bool SegmentsIntersect(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd) =>
        SegmentsIntersect(
            (firstStart.X, firstStart.Y),
            (firstEnd.X, firstEnd.Y),
            (secondStart.X, secondStart.Y),
            (secondEnd.X, secondEnd.Y));

    /// <summary>
    /// Determines whether a line segment intersects a rectangle, including boundary touches.
    /// </summary>
    /// <param name="start">Segment start point.</param>
    /// <param name="end">Segment end point.</param>
    /// <param name="left">Rectangle left X coordinate.</param>
    /// <param name="bottom">Rectangle lower Y coordinate.</param>
    /// <param name="right">Rectangle right X coordinate.</param>
    /// <param name="top">Rectangle upper Y coordinate.</param>
    public static bool SegmentIntersectsRectangle(
        (double X, double Y) start,
        (double X, double Y) end,
        double left,
        double bottom,
        double right,
        double top) {
        double minX = Math.Min(left, right);
        double maxX = Math.Max(left, right);
        double minY = Math.Min(bottom, top);
        double maxY = Math.Max(bottom, top);

        if (Math.Max(start.X, end.X) < minX - GeometryTolerance ||
            Math.Min(start.X, end.X) > maxX + GeometryTolerance ||
            Math.Max(start.Y, end.Y) < minY - GeometryTolerance ||
            Math.Min(start.Y, end.Y) > maxY + GeometryTolerance) {
            return false;
        }

        if (PointInsideRectangle(start, minX, minY, maxX, maxY) ||
            PointInsideRectangle(end, minX, minY, maxX, maxY)) {
            return true;
        }

        (double X, double Y) bottomLeft = (minX, minY);
        (double X, double Y) bottomRight = (maxX, minY);
        (double X, double Y) topRight = (maxX, maxY);
        (double X, double Y) topLeft = (minX, maxY);
        return SegmentsIntersect(start, end, bottomLeft, bottomRight) ||
               SegmentsIntersect(start, end, bottomRight, topRight) ||
               SegmentsIntersect(start, end, topRight, topLeft) ||
               SegmentsIntersect(start, end, topLeft, bottomLeft);
    }

    /// <summary>
    /// Determines whether a line segment intersects a rectangle, including boundary touches.
    /// </summary>
    /// <param name="start">Segment start point.</param>
    /// <param name="end">Segment end point.</param>
    /// <param name="left">Rectangle left X coordinate.</param>
    /// <param name="bottom">Rectangle lower Y coordinate.</param>
    /// <param name="right">Rectangle right X coordinate.</param>
    /// <param name="top">Rectangle upper Y coordinate.</param>
    public static bool SegmentIntersectsRectangle(
        OfficePoint start,
        OfficePoint end,
        double left,
        double bottom,
        double right,
        double top) =>
        SegmentIntersectsRectangle(
            (start.X, start.Y),
            (end.X, end.Y),
            left,
            bottom,
            right,
            top);

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
    /// Builds a connector polyline from endpoints, optional waypoints, and optional right-angle fallback routing.
    /// </summary>
    /// <param name="start">Connector start point.</param>
    /// <param name="end">Connector end point.</param>
    /// <param name="waypoints">Explicit intermediate points. When present, these take precedence over right-angle fallback routing.</param>
    /// <param name="useRightAngleFallback">When true and no explicit waypoints exist, inserts one orthogonal elbow at start X and end Y.</param>
    /// <returns>A connector polyline containing start, any intermediate points, and end.</returns>
    public static List<OfficePoint> BuildConnectorPolyline(
        OfficePoint start,
        OfficePoint end,
        IReadOnlyList<OfficePoint>? waypoints,
        bool useRightAngleFallback) {
        List<OfficePoint> points = new() { start };
        if (waypoints != null && waypoints.Count > 0) {
            for (int i = 0; i < waypoints.Count; i++) {
                points.Add(waypoints[i]);
            }
        } else if (useRightAngleFallback) {
            points.Add(new OfficePoint(start.X, end.Y));
        }

        points.Add(end);
        return points;
    }

    /// <summary>
    /// Builds a tuple connector polyline from endpoints, optional waypoints, and optional right-angle fallback routing.
    /// </summary>
    /// <param name="start">Connector start point.</param>
    /// <param name="end">Connector end point.</param>
    /// <param name="waypoints">Explicit intermediate points. When present, these take precedence over right-angle fallback routing.</param>
    /// <param name="useRightAngleFallback">When true and no explicit waypoints exist, inserts one orthogonal elbow at start X and end Y.</param>
    /// <returns>A connector polyline containing start, any intermediate points, and end.</returns>
    public static List<(double X, double Y)> BuildConnectorPolyline(
        (double X, double Y) start,
        (double X, double Y) end,
        IReadOnlyList<(double X, double Y)>? waypoints,
        bool useRightAngleFallback) {
        List<(double X, double Y)> points = new() { start };
        if (waypoints != null && waypoints.Count > 0) {
            for (int i = 0; i < waypoints.Count; i++) {
                points.Add(waypoints[i]);
            }
        } else if (useRightAngleFallback) {
            points.Add((start.X, end.Y));
        }

        points.Add(end);
        return points;
    }

    /// <summary>
    /// Samples a quadratic Bezier curve into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="start">Curve start point.</param>
    /// <param name="control">Quadratic control point.</param>
    /// <param name="end">Curve end point.</param>
    /// <param name="segments">Number of line segments used to approximate the curve.</param>
    public static List<OfficePoint> CreateQuadraticBezierPoints(OfficePoint start, OfficePoint control, OfficePoint end, int segments) {
        EnsurePositiveSegmentCount(segments);
        var points = new List<OfficePoint>(segments);
        for (int i = 1; i <= segments; i++) {
            double t = i / (double)segments;
            double inverse = 1D - t;
            double x = (inverse * inverse * start.X) + (2D * inverse * t * control.X) + (t * t * end.X);
            double y = (inverse * inverse * start.Y) + (2D * inverse * t * control.Y) + (t * t * end.Y);
            points.Add(new OfficePoint(x, y));
        }

        return points;
    }

    /// <summary>
    /// Samples a tuple quadratic Bezier curve into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="start">Curve start point.</param>
    /// <param name="control">Quadratic control point.</param>
    /// <param name="end">Curve end point.</param>
    /// <param name="segments">Number of line segments used to approximate the curve.</param>
    public static List<(double X, double Y)> CreateQuadraticBezierPoints(
        (double X, double Y) start,
        (double X, double Y) control,
        (double X, double Y) end,
        int segments) {
        List<OfficePoint> sampled = CreateQuadraticBezierPoints(
            new OfficePoint(start.X, start.Y),
            new OfficePoint(control.X, control.Y),
            new OfficePoint(end.X, end.Y),
            segments);
        var points = new List<(double X, double Y)>(sampled.Count);
        for (int i = 0; i < sampled.Count; i++) {
            points.Add((sampled[i].X, sampled[i].Y));
        }

        return points;
    }

    /// <summary>
    /// Samples a cubic Bezier curve into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="start">Curve start point.</param>
    /// <param name="control1">First cubic control point.</param>
    /// <param name="control2">Second cubic control point.</param>
    /// <param name="end">Curve end point.</param>
    /// <param name="segments">Number of line segments used to approximate the curve.</param>
    public static List<OfficePoint> CreateCubicBezierPoints(OfficePoint start, OfficePoint control1, OfficePoint control2, OfficePoint end, int segments) {
        EnsurePositiveSegmentCount(segments);
        var points = new List<OfficePoint>(segments);
        for (int i = 1; i <= segments; i++) {
            double t = i / (double)segments;
            double inverse = 1D - t;
            double inverseSquared = inverse * inverse;
            double tSquared = t * t;
            double x = (inverseSquared * inverse * start.X) +
                       (3D * inverseSquared * t * control1.X) +
                       (3D * inverse * tSquared * control2.X) +
                       (tSquared * t * end.X);
            double y = (inverseSquared * inverse * start.Y) +
                       (3D * inverseSquared * t * control1.Y) +
                       (3D * inverse * tSquared * control2.Y) +
                       (tSquared * t * end.Y);
            points.Add(new OfficePoint(x, y));
        }

        return points;
    }

    /// <summary>
    /// Samples a tuple cubic Bezier curve into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="start">Curve start point.</param>
    /// <param name="control1">First cubic control point.</param>
    /// <param name="control2">Second cubic control point.</param>
    /// <param name="end">Curve end point.</param>
    /// <param name="segments">Number of line segments used to approximate the curve.</param>
    public static List<(double X, double Y)> CreateCubicBezierPoints(
        (double X, double Y) start,
        (double X, double Y) control1,
        (double X, double Y) control2,
        (double X, double Y) end,
        int segments) {
        List<OfficePoint> sampled = CreateCubicBezierPoints(
            new OfficePoint(start.X, start.Y),
            new OfficePoint(control1.X, control1.Y),
            new OfficePoint(control2.X, control2.Y),
            new OfficePoint(end.X, end.Y),
            segments);
        var points = new List<(double X, double Y)>(sampled.Count);
        for (int i = 0; i < sampled.Count; i++) {
            points.Add((sampled[i].X, sampled[i].Y));
        }

        return points;
    }

    /// <summary>
    /// Samples an elliptical arc into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="centerX">Arc center X coordinate.</param>
    /// <param name="centerY">Arc center Y coordinate.</param>
    /// <param name="radiusX">Horizontal radius before optional rotation.</param>
    /// <param name="radiusY">Vertical radius before optional rotation.</param>
    /// <param name="startRadians">Start angle in radians.</param>
    /// <param name="sweepRadians">Sweep angle in radians.</param>
    /// <param name="segments">Number of line segments used to approximate the arc.</param>
    /// <param name="rotationRadians">Optional rotation angle in radians.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    public static List<OfficePoint> CreateEllipticalArcPoints(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        double startRadians,
        double sweepRadians,
        int segments,
        double rotationRadians = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D) {
        EnsurePositiveSegmentCount(segments);
        if (!IsFinite(radiusX) || radiusX <= 0D || !IsFinite(radiusY) || radiusY <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(radiusX), "Arc radii must be positive finite values.");
        }

        var points = new List<OfficePoint>(segments);
        for (int i = 1; i <= segments; i++) {
            double t = i / (double)segments;
            double angle = startRadians + (sweepRadians * t);
            OfficePoint point = new(
                centerX + (Math.Cos(angle) * radiusX),
                centerY + (Math.Sin(angle) * radiusY));
            points.Add(Math.Abs(rotationRadians) > 0.000001D
                ? RotatePoint(point, rotationCenterX, rotationCenterY, rotationRadians)
                : point);
        }

        return points;
    }

    /// <summary>
    /// Samples a tuple elliptical arc into straight-line points, excluding the start point and including the end point.
    /// </summary>
    /// <param name="centerX">Arc center X coordinate.</param>
    /// <param name="centerY">Arc center Y coordinate.</param>
    /// <param name="radiusX">Horizontal radius before optional rotation.</param>
    /// <param name="radiusY">Vertical radius before optional rotation.</param>
    /// <param name="startRadians">Start angle in radians.</param>
    /// <param name="sweepRadians">Sweep angle in radians.</param>
    /// <param name="segments">Number of line segments used to approximate the arc.</param>
    /// <param name="rotationRadians">Optional rotation angle in radians.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    public static List<(double X, double Y)> CreateEllipticalArcPointsAsTuples(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        double startRadians,
        double sweepRadians,
        int segments,
        double rotationRadians = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D) {
        List<OfficePoint> sampled = CreateEllipticalArcPoints(
            centerX,
            centerY,
            radiusX,
            radiusY,
            startRadians,
            sweepRadians,
            segments,
            rotationRadians,
            rotationCenterX,
            rotationCenterY);
        var points = new List<(double X, double Y)>(sampled.Count);
        for (int i = 0; i < sampled.Count; i++) {
            points.Add((sampled[i].X, sampled[i].Y));
        }

        return points;
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

    private static void EnsurePositiveSegmentCount(int segments) {
        if (segments <= 0) {
            throw new ArgumentOutOfRangeException(nameof(segments), "Curve segment count must be positive.");
        }
    }

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && IsFinite(value) ? value : 0D;

    private static bool IsFinite(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value);

    private static double Orientation((double X, double Y) a, (double X, double Y) b, (double X, double Y) c) =>
        ((b.X - a.X) * (c.Y - a.Y)) - ((b.Y - a.Y) * (c.X - a.X));

    private static bool OnSegment((double X, double Y) a, (double X, double Y) b, (double X, double Y) c) =>
        b.X >= Math.Min(a.X, c.X) - GeometryTolerance &&
        b.X <= Math.Max(a.X, c.X) + GeometryTolerance &&
        b.Y >= Math.Min(a.Y, c.Y) - GeometryTolerance &&
        b.Y <= Math.Max(a.Y, c.Y) + GeometryTolerance;

    private static bool PointInsideRectangle((double X, double Y) point, double minX, double minY, double maxX, double maxY) =>
        point.X >= minX - GeometryTolerance &&
        point.X <= maxX + GeometryTolerance &&
        point.Y >= minY - GeometryTolerance &&
        point.Y <= maxY + GeometryTolerance;

    private static bool IsZero(double value) =>
        Math.Abs(value) < GeometryTolerance;
}
