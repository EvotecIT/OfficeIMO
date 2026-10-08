using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

[Flags]
internal enum OfficeConnectorDirections {
    Left = 1, Right = 2, Up = 4, Down = 8,
    Horizontal = Left | Right, Vertical = Up | Down, Any = Horizontal | Vertical
}

public static partial class OfficeGeometry {
    // Try the existing short routes first. Exit stubs add five-segment candidates when both
    // endpoints point away from each other, without multiplying the two-dimensional lane grid.
    internal static IEnumerable<OfficePoint[]> EnumerateConstrainedOrthogonalConnectorRoutes(OfficePoint start, OfficePoint end,
        OfficeConnectorDirections startDirections, OfficeConnectorDirections endDirections, double step, int maxLanes, double exitClearance,
        (double Left, double Top, double Right, double Bottom)? startBounds,
        (double Left, double Top, double Right, double Bottom)? endBounds) {
        foreach (OfficePoint[] route in EnumerateOrthogonalConnectorRoutes(start, end, step, maxLanes)) {
            OfficePoint[]? normalized = NormalizeOrthogonalRoute(route);
            if (normalized != null && MatchesConnectorDirections(normalized, startDirections, endDirections)) yield return normalized;
        }
        foreach (OfficeConnectorDirections startDirection in CardinalDirections(startDirections))
            foreach (OfficeConnectorDirections endDirection in CardinalDirections(endDirections)) {
                OfficePoint startExit = ConnectorExit(start, startDirection, step, exitClearance, startBounds);
                OfficePoint endExit = ConnectorExit(end, endDirection, step, exitClearance, endBounds);
                foreach (double offset in ConnectorLaneOffsets(step, maxLanes))
                    foreach (OfficePoint[] route in ConnectorExitRoutes(start, end, startExit, endExit, offset, startDirections, endDirections)) yield return route;
                // Extend both departures together and couple the bridge displacement to that extension.
                // This permits detours past intervening obstacles without a Cartesian grid of stub lengths.
                for (int lane = 1; lane <= maxLanes; lane++) {
                    double extension = lane * step;
                    OfficePoint extendedStart = ConnectorExit(startExit, startDirection, extension, 0, null);
                    OfficePoint extendedEnd = ConnectorExit(endExit, endDirection, extension, 0, null);
                    foreach (double offset in new[] { 0D, extension, -extension })
                        foreach (OfficePoint[] route in ConnectorExitRoutes(start, end, extendedStart, extendedEnd, offset, startDirections, endDirections)) yield return route;
                }
            }
    }

    private static IEnumerable<OfficePoint[]> ConnectorExitRoutes(OfficePoint start, OfficePoint end, OfficePoint startExit, OfficePoint endExit,
        double offset, OfficeConnectorDirections startDirections, OfficeConnectorDirections endDirections) {
        foreach (bool horizontal in new[] { true, false }) {
            OfficePoint[] bridge = CreateOrthogonalConnectorRoute(startExit, endExit, horizontal, offset);
            var route = new[] { start, bridge[0], bridge[1], bridge[2], bridge[3], end };
            OfficePoint[]? normalized = NormalizeOrthogonalRoute(route);
            if (normalized != null && MatchesConnectorDirections(normalized, startDirections, endDirections)) yield return normalized;
        }
    }

    internal static bool MatchesConnectorDirections(IReadOnlyList<OfficePoint> route, OfficeConnectorDirections start,
        OfficeConnectorDirections end, double axisTolerance = 1e-9) {
        if (route.Count < 2) return false;
        for (int i = 1; i < route.Count; i++) {
            double x = route[i].X - route[i - 1].X, y = route[i].Y - route[i - 1].Y;
            if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y) ||
                (Math.Abs(x) <= 1e-9 && Math.Abs(y) <= 1e-9) || Math.Min(Math.Abs(x), Math.Abs(y)) > axisTolerance) return false;
        }
        return (start & ConnectorDirection(route[0], route[1])) != 0 &&
            (end & ConnectorDirection(route[route.Count - 1], route[route.Count - 2])) != 0;
    }

    private static OfficeConnectorDirections ConnectorDirection(OfficePoint from, OfficePoint to) =>
        Math.Abs(to.X - from.X) >= Math.Abs(to.Y - from.Y)
            ? to.X > from.X ? OfficeConnectorDirections.Right : OfficeConnectorDirections.Left
            : to.Y > from.Y ? OfficeConnectorDirections.Down : OfficeConnectorDirections.Up;

    private static IEnumerable<OfficeConnectorDirections> CardinalDirections(OfficeConnectorDirections directions) {
        foreach (OfficeConnectorDirections direction in new[] { OfficeConnectorDirections.Left, OfficeConnectorDirections.Right,
            OfficeConnectorDirections.Up, OfficeConnectorDirections.Down }) if ((directions & direction) != 0) yield return direction;
    }

    private static OfficePoint ConnectorExit(OfficePoint point, OfficeConnectorDirections direction, double step, double clearance,
        (double Left, double Top, double Right, double Bottom)? bounds) {
        double distance = Math.Max(step, 0.15);
        if (bounds is { } box) {
            double edgeDistance = direction switch {
                OfficeConnectorDirections.Left => point.X - box.Left,
                OfficeConnectorDirections.Right => box.Right - point.X,
                OfficeConnectorDirections.Up => point.Y - box.Top,
                _ => box.Bottom - point.Y
            };
            distance = Math.Max(distance, edgeDistance + clearance);
        }
        OfficePoint result = direction switch {
            OfficeConnectorDirections.Left => new OfficePoint(point.X - distance, point.Y),
            OfficeConnectorDirections.Right => new OfficePoint(point.X + distance, point.Y),
            OfficeConnectorDirections.Up => new OfficePoint(point.X, point.Y - distance),
            _ => new OfficePoint(point.X, point.Y + distance)
        };
        ValidateRoutingPoint(result); return result;
    }

    private static OfficePoint[]? NormalizeOrthogonalRoute(IReadOnlyList<OfficePoint> route) {
        var result = new List<OfficePoint>();
        foreach (OfficePoint point in route) {
            if (result.Count > 0 && SameRoutingPoint(result[result.Count - 1], point)) continue;
            if (result.Count > 1) {
                OfficePoint before = result[result.Count - 2], last = result[result.Count - 1];
                double ax = last.X - before.X, ay = last.Y - before.Y, bx = point.X - last.X, by = point.Y - last.Y;
                if ((Math.Abs(ax) <= 1e-9 && Math.Abs(bx) <= 1e-9) || (Math.Abs(ay) <= 1e-9 && Math.Abs(by) <= 1e-9)) {
                    if (ax * bx + ay * by < 0) return null; // Do not reverse immediately over an exit segment.
                    result.RemoveAt(result.Count - 1);
                }
            }
            result.Add(point);
        }
        return result.Count < 2 ? null : result.ToArray();
    }
    private static bool SameRoutingPoint(OfficePoint a, OfficePoint b) => Math.Abs(a.X - b.X) <= 1e-9 && Math.Abs(a.Y - b.Y) <= 1e-9;
}
