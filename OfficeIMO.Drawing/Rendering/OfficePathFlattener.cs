using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal sealed class OfficeFlattenedPathContour {
    internal OfficeFlattenedPathContour(IReadOnlyList<OfficePoint> points, bool closed) {
        Points = points ?? throw new ArgumentNullException(nameof(points));
        Closed = closed;
    }

    internal IReadOnlyList<OfficePoint> Points { get; }

    internal bool Closed { get; }
}

internal static class OfficePathFlattener {
    private const int DefaultCurveSegments = 24;

    internal static IReadOnlyList<OfficeFlattenedPathContour> Flatten(
        IReadOnlyList<OfficePathCommand> commands,
        double offsetX,
        double offsetY,
        double scale,
        int curveSegments = DefaultCurveSegments) {
        if (commands == null) {
            throw new ArgumentNullException(nameof(commands));
        }

        if (curveSegments <= 0) {
            throw new ArgumentOutOfRangeException(nameof(curveSegments), "Curve segment count must be positive.");
        }

        var contours = new List<OfficeFlattenedPathContour>();
        List<OfficePoint>? current = null;
        OfficePoint currentPoint = default;
        bool hasCurrentPoint = false;

        foreach (OfficePathCommand command in commands) {
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    AddOpenContour(contours, current);
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    current = new List<OfficePoint> { currentPoint };
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.LineTo:
                    EnsureCurrentContour(ref current, currentPoint, hasCurrentPoint);
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    current!.Add(currentPoint);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    EnsureCurrentContour(ref current, currentPoint, hasCurrentPoint);
                    AddQuadraticPoints(
                        current!,
                        currentPoint,
                        Transform(command.ControlPoint1, offsetX, offsetY, scale),
                        Transform(command.Point, offsetX, offsetY, scale),
                        curveSegments);
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    EnsureCurrentContour(ref current, currentPoint, hasCurrentPoint);
                    AddCubicPoints(
                        current!,
                        currentPoint,
                        Transform(command.ControlPoint1, offsetX, offsetY, scale),
                        Transform(command.ControlPoint2, offsetX, offsetY, scale),
                        Transform(command.Point, offsetX, offsetY, scale),
                        curveSegments);
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.Close:
                    AddClosedContour(contours, current);
                    current = null;
                    hasCurrentPoint = false;
                    break;
            }
        }

        AddOpenContour(contours, current);
        return contours;
    }

    private static void EnsureCurrentContour(ref List<OfficePoint>? current, OfficePoint currentPoint, bool hasCurrentPoint) {
        if (current == null) {
            current = hasCurrentPoint ? new List<OfficePoint> { currentPoint } : new List<OfficePoint>();
        }
    }

    private static void AddOpenContour(List<OfficeFlattenedPathContour> contours, List<OfficePoint>? points) {
        if (points != null && points.Count >= 2) {
            contours.Add(new OfficeFlattenedPathContour(points, closed: false));
        }
    }

    private static void AddClosedContour(List<OfficeFlattenedPathContour> contours, List<OfficePoint>? points) {
        if (points != null && points.Count >= 2) {
            contours.Add(new OfficeFlattenedPathContour(points, closed: true));
        }
    }

    private static void AddQuadraticPoints(List<OfficePoint> points, OfficePoint start, OfficePoint control, OfficePoint end, int segments) {
        for (int i = 1; i <= segments; i++) {
            double t = i / (double)segments;
            double inverse = 1D - t;
            double x = (inverse * inverse * start.X) + (2D * inverse * t * control.X) + (t * t * end.X);
            double y = (inverse * inverse * start.Y) + (2D * inverse * t * control.Y) + (t * t * end.Y);
            points.Add(new OfficePoint(x, y));
        }
    }

    private static void AddCubicPoints(List<OfficePoint> points, OfficePoint start, OfficePoint control1, OfficePoint control2, OfficePoint end, int segments) {
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
    }

    private static OfficePoint Transform(OfficePoint point, double offsetX, double offsetY, double scale) =>
        new OfficePoint(offsetX + (point.X * scale), offsetY + (point.Y * scale));
}
