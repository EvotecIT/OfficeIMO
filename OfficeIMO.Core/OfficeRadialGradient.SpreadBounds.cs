using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRadialGradient {
    // A finite convex painted region bounds native repeated fields without changing
    // their first-intersection semantics or manufacturing infinitely many stops.
    internal bool TryGetSpreadCycles(IEnumerable<OfficePoint> corners, bool nativeFirstIntersection,
        bool reflect, int maximumCycles, out int cycles) {
        cycles = 0;
        // Interior point foci have a convex gauge. Native exterior foci instead
        // use the first physical root, bounded by the square root of the root
        // product C/A. C is convex, so its maximum also occurs at a box corner.
        // Boundary foci have A=0. C/(-B) is convex where B<0, so a
        // box strictly inside that half-plane also has a finite corner bound.
        if (StartRadiusX != 0D || StartRadiusY != 0D) return false;
        double dx = (StartX - EndX) / EndRadiusX;
        double dy = (StartY - EndY) / EndRadiusY;
        double a = dx * dx + dy * dy - 1D;
        bool exterior = a > 0D, boundary = a == 0D;
        if (a >= 0D && !nativeFirstIntersection) return false;
        double maximum = 1D, halfPlaneMaximum = 1D;
        foreach (var point in corners) {
            double bound;
            if (exterior || boundary) {
                var local = CoordinateTransform.Invert().TransformPoint(point);
                double px = (local.X - StartX) / EndRadiusX;
                double py = (local.Y - StartY) / EndRadiusY;
                double b = 2D * (px * dx + py * dy);
                // For exterior roots, t=2C/(-B+sqrt(D)) <= 2C/(-B).
                // This convex bound stays useful as A approaches zero.
                halfPlaneMaximum = b < 0D
                    ? Math.Max(halfPlaneMaximum, -2D * (px * px + py * py) / b)
                    : double.PositiveInfinity;
                if (boundary) {
                    if (b >= 0D) return false;
                    bound = -(px * px + py * py) / b;
                } else bound = Math.Sqrt((px * px + py * py) / a);
            } else bound = SampleUnboundedRatio(point.X, point.Y);
            maximum = Math.Max(maximum, bound);
        }
        if (exterior) maximum = Math.Min(maximum, halfPlaneMaximum);
        maximum = Math.Ceiling(maximum);
        // Native Reflect paints offset zero outside the cone. Ending on an
        // even cycle makes that color the reversed field's exterior endpoint.
        if ((exterior || boundary) && reflect && maximum % 2D != 0D) maximum++;
        if (double.IsNaN(maximum) || double.IsInfinity(maximum) || maximum > maximumCycles) return false;
        cycles = (int)maximum;
        return true;
    }
}
