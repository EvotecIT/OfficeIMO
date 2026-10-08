using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal sealed class OfficeFlattenedPathContour {
    internal OfficeFlattenedPathContour(IReadOnlyList<OfficePoint> points, bool closed) : this(points, closed, true, true) { }

    internal OfficeFlattenedPathContour(IReadOnlyList<OfficePoint> points, bool closed, bool useStartLineCap, bool useEndLineCap) {
        Points = points ?? throw new ArgumentNullException(nameof(points));
        Closed = closed;
        UseStartLineCap = useStartLineCap; UseEndLineCap = useEndLineCap;
    }

    internal IReadOnlyList<OfficePoint> Points { get; }

    internal bool Closed { get; }
    internal bool UseStartLineCap { get; }
    internal bool UseEndLineCap { get; }
}

internal static class OfficePathFlattener {
    private const int MaximumFlattenedPoints = 1_000_000;

    internal static IReadOnlyList<OfficeFlattenedPathContour> Flatten(
        IReadOnlyList<OfficePathCommand> commands,
        double offsetX,
        double offsetY,
        double scale,
        int curveSegments = 0,
        double pixelsPerUnit = 1D) => FlattenCore(commands, offsetX, offsetY, scale, curveSegments, pixelsPerUnit, false);

    internal static IReadOnlyList<OfficeFlattenedPathContour> FlattenNativeStroke(IReadOnlyList<OfficePathCommand> commands,
        double pixelsPerUnit) => FlattenCore(commands, 0D, 0D, 1D, 0, pixelsPerUnit, true);

    private static IReadOnlyList<OfficeFlattenedPathContour> FlattenCore(IReadOnlyList<OfficePathCommand> commands,
        double offsetX, double offsetY, double scale, int curveSegments, double pixelsPerUnit,
        bool retainSinglePointClosedContours) {
        if (commands == null) {
            throw new ArgumentNullException(nameof(commands));
        }

        if (curveSegments < 0) {
            throw new ArgumentOutOfRangeException(nameof(curveSegments), "Curve segment count must be non-negative; zero selects adaptive flattening.");
        }

        var contours = new List<OfficeFlattenedPathContour>();
        List<OfficePoint>? current = null;
        OfficePoint currentPoint = default;
        bool hasCurrentPoint = false;

        int pointCount = 0;
        foreach (OfficePathCommand command in commands) {
            int previousCount = current?.Count ?? 0;
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
                    current!.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(
                        currentPoint,
                        Transform(command.ControlPoint1, offsetX, offsetY, scale),
                        Transform(command.Point, offsetX, offsetY, scale),
                        curveSegments > 0 ? curveSegments : OfficeCurveFlattening.QuadraticSegments(currentPoint,
                            Transform(command.ControlPoint1, offsetX, offsetY, scale), Transform(command.Point, offsetX, offsetY, scale), pixelsPerUnit)));
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    EnsureCurrentContour(ref current, currentPoint, hasCurrentPoint);
                    current!.AddRange(OfficeGeometry.CreateCubicBezierPoints(
                        currentPoint,
                        Transform(command.ControlPoint1, offsetX, offsetY, scale),
                        Transform(command.ControlPoint2, offsetX, offsetY, scale),
                        Transform(command.Point, offsetX, offsetY, scale),
                        curveSegments > 0 ? curveSegments : OfficeCurveFlattening.CubicSegments(currentPoint,
                            Transform(command.ControlPoint1, offsetX, offsetY, scale), Transform(command.ControlPoint2, offsetX, offsetY, scale),
                            Transform(command.Point, offsetX, offsetY, scale), pixelsPerUnit)));
                    currentPoint = Transform(command.Point, offsetX, offsetY, scale);
                    hasCurrentPoint = true;
                    break;
                case OfficePathCommandKind.Close:
                    AddClosedContour(contours, current, retainSinglePointClosedContours);
                    current = null;
                    hasCurrentPoint = false;
                    break;
            }
            pointCount += Math.Max(1, (current?.Count ?? 0) - previousCount);
            if (pointCount > MaximumFlattenedPoints) throw new InvalidOperationException("Flattened path exceeds the point limit.");
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

    private static void AddClosedContour(List<OfficeFlattenedPathContour> contours, List<OfficePoint>? points, bool retainSinglePoint) {
        if (points != null && (points.Count >= 2 || retainSinglePoint && points.Count == 1)) {
            contours.Add(new OfficeFlattenedPathContour(points, closed: true));
        }
    }

    private static OfficePoint Transform(OfficePoint point, double offsetX, double offsetY, double scale) =>
        new OfficePoint(offsetX + (point.X * scale), offsetY + (point.Y * scale));
}
