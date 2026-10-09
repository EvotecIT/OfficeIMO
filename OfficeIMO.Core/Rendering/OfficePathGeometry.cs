using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Geometry queries over normalized paths, without replacing curves by control-point hulls.</summary>
internal static class OfficePathGeometry {
    internal static (double Left, double Top, double Right, double Bottom) Bounds(IReadOnlyList<OfficePathCommand> commands) {
        double left = double.PositiveInfinity, top = left, right = double.NegativeInfinity, bottom = right;
        OfficePoint current = default, start = default;
        bool hasSegment = false;
        void Include(OfficePoint p) { left = Math.Min(left, p.X); top = Math.Min(top, p.Y); right = Math.Max(right, p.X); bottom = Math.Max(bottom, p.Y); }
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) { current = start = command.Point; hasSegment = false; continue; }
            if (command.Kind == OfficePathCommandKind.Close) { if (hasSegment) { Include(current); Include(start); } current = start; continue; }
            Include(current); Include(command.Point);
            hasSegment = true;
            if (command.Kind is OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo) {
                OfficePoint origin = current;
                bool cubic = command.Kind == OfficePathCommandKind.CubicBezierTo;
                foreach (bool horizontal in new[] { true, false }) {
                    double p0 = horizontal ? origin.X : origin.Y, p1 = horizontal ? command.ControlPoint1.X : command.ControlPoint1.Y;
                    double p2 = horizontal ? (cubic ? command.ControlPoint2.X : command.Point.X) : (cubic ? command.ControlPoint2.Y : command.Point.Y);
                    double p3 = horizontal ? command.Point.X : command.Point.Y;
                    foreach (double t in Extrema(p0, p1, p2, p3, cubic)) {
                        double u = 1 - t;
                        Include(cubic ? new OfficePoint(u*u*u*origin.X + 3*u*u*t*command.ControlPoint1.X + 3*u*t*t*command.ControlPoint2.X + t*t*t*command.Point.X,
                            u*u*u*origin.Y + 3*u*u*t*command.ControlPoint1.Y + 3*u*t*t*command.ControlPoint2.Y + t*t*t*command.Point.Y) :
                            new OfficePoint(u*u*origin.X + 2*u*t*command.ControlPoint1.X + t*t*command.Point.X,
                                u*u*origin.Y + 2*u*t*command.ControlPoint1.Y + t*t*command.Point.Y));
                    }
                }
            }
            current = command.Point;
        }
        return (left, top, right, bottom);
    }

    private static IEnumerable<double> Extrema(double p0, double p1, double p2, double p3, bool cubic) {
        // Normalize before polynomial arithmetic so finite large coordinates do not overflow.
        double scale = Math.Max(Math.Max(Math.Abs(p0), Math.Abs(p1)), Math.Max(Math.Abs(p2), Math.Abs(p3)));
        if (scale == 0) yield break;
        p0 /= scale; p1 /= scale; p2 /= scale; p3 /= scale;
        double a = cubic ? -p0 + 3*p1 - 3*p2 + p3 : 0;
        double b = cubic ? 2*(p0 - 2*p1 + p2) : p0 - 2*p1 + p2;
        double c = p1 - p0;
        if (a == 0) {
            if (b != 0) { double t = -c/b; if (t > 0 && t < 1) yield return t; }
            yield break;
        }
        double discriminant = b*b - 4*a*c;
        if (discriminant < 0) yield break;
        double root = Math.Sqrt(discriminant), q = -0.5*(b + (b >= 0 ? root : -root));
        double first = q/a;
        if (first > 0 && first < 1) yield return first;
        if (q != 0) { double second = c/q; if (second > 0 && second < 1) yield return second; }
    }

    /// <summary>Gets exact endpoint tangent directions for one open contour, skipping zero-length segments.</summary>
    internal static bool TryOpenEndpoints(IReadOnlyList<OfficePathCommand> commands, out OfficePoint start, out OfficePoint startInward, out OfficePoint end, out OfficePoint endInward) {
        start = end = startInward = endInward = default;
        bool moved = false, hasSegment = false;
        OfficePoint current = default;
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) {
                if (moved) return false;
                current = start = command.Point; moved = true; continue;
            }
            if (command.Kind == OfficePathCommandKind.Close || !moved) return false;
            var first = command.Kind is OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo ? command.ControlPoint1 : command.Point;
            var last = command.Kind == OfficePathCommandKind.CubicBezierTo ? command.ControlPoint2 : first;
            OfficePoint forward = Difference(first, current);
            if (Zero(forward) && command.Kind == OfficePathCommandKind.CubicBezierTo) forward = Difference(last, current);
            if (Zero(forward)) forward = Difference(command.Point, current);
            OfficePoint backward = Difference(last, command.Point);
            if (Zero(backward) && command.Kind == OfficePathCommandKind.CubicBezierTo) backward = Difference(first, command.Point);
            if (Zero(backward)) backward = Difference(current, command.Point);
            if (!hasSegment && !Zero(forward)) { startInward = forward; hasSegment = true; }
            if (!Zero(backward)) endInward = backward;
            current = end = command.Point;
        }
        return hasSegment && !Zero(endInward);
    }
    private static OfficePoint Difference(OfficePoint a, OfficePoint b) => new OfficePoint(a.X - b.X, a.Y - b.Y);
    private static bool Zero(OfficePoint p) => p.X == 0 && p.Y == 0;
}
