using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRadialGradient {
    private OfficeTransform _inverseCoordinates = OfficeTransform.Identity;
    /// <summary>Affine mapping from gradient coordinates to normalized shape coordinates.</summary>
    public OfficeTransform CoordinateTransform { get; private set; } = OfficeTransform.Identity;

    /// <summary>Returns a gradient transformed in normalized shape coordinates, preserving its circular or elliptical field.</summary>
    /// <exception cref="ArgumentException">The combined transform cannot be inverted.</exception>
    public OfficeRadialGradient TransformCoordinates(OfficeTransform transform) {
        var combined = CoordinateTransform.Then(transform);
        if (!combined.TryInvert(out var inverse)) throw new ArgumentException("Radial gradient coordinate transform must be invertible.", nameof(transform));
        var copy = Clone(); copy.CoordinateTransform = combined; copy._inverseCoordinates = inverse;
        return copy;
    }

    internal OfficeRadialGradient WithStops(IReadOnlyList<OfficeGradientStop> stops) {
        var copy = Clone(); copy.Stops = ValidateStops(stops); return copy;
    }

    // No physical circle paints an ordinary SVG/PDF field outside its cone.
    // Native XPS adapters explicitly retain their endpoint paint there.
    private double MissingCircleRatio => _paintOutsideCone ? 0D : double.NaN;

    internal double SampleRatio(double x, double y) => Math.Max(0D, Math.Min(1D, SampleUnboundedRatio(x, y)));

    internal double SampleUnboundedRatio(double x, double y) {
        var point = _inverseCoordinates.TransformPoint(new OfficePoint(x, y));
        x = point.X; y = point.Y;
        double endRadiusX = Math.Max(Math.Max(StartRadiusX, EndRadiusX), 0.0000001D);
        double endRadiusY = Math.Max(Math.Max(StartRadiusY, EndRadiusY), 0.0000001D);
        // Solve from a point end using u=1-t. Subtracting unit-radius squares
        // near t=1 can round an outside-cone root onto a physical zero radius.
        bool pointEnd = EndRadiusX == 0D && EndRadiusY == 0D;
        double startRadius = pointEnd ? 0D : StartRadiusX / endRadiusX;
        double vx = (x - (pointEnd ? EndX : StartX)) / endRadiusX;
        double vy = (y - (pointEnd ? EndY : StartY)) / endRadiusY;
        double dx = (EndX - StartX) / endRadiusX * (pointEnd ? -1D : 1D);
        double dy = (EndY - StartY) / endRadiusY * (pointEnd ? -1D : 1D);
        double dr = (EndRadiusX - StartRadiusX) / endRadiusX * (pointEnd ? -1D : 1D);
        double a = (dx * dx) + (dy * dy) - (dr * dr);
        double b = -2D * ((vx * dx) + (vy * dy) + (startRadius * dr));
        double c = (vx * vx) + (vy * vy) - (startRadius * startRadius);
        // Small coefficients still describe physical circles. An absolute
        // threshold erases near-boundary fields after a large coordinate map.
        if (a == 0D) {
            if (b == 0D) {
                if (c != 0D) return MissingCircleRatio;
                // At the common tangent point every physical circle intersects.
                // Select the latest one, including the zero-radius point end.
                return pointEnd ? 1D : dr < 0D ? -startRadius / dr : double.PositiveInfinity;
            }

            double linearRatio = -c / b;
            return startRadius + linearRatio * dr >= 0D ? (pointEnd ? 1D - linearRatio : linearRatio) : MissingCircleRatio;
        }

        double discriminant = (b * b) - (4D * a * c);
        if (discriminant < 0D) {
            return MissingCircleRatio;
        }

        double sqrt = Math.Sqrt(discriminant);
        // Compute the non-cancelling numerator first and recover the other
        // root from their product. This also preserves almost-linear fields.
        double q = -0.5D * (b + (b >= 0D ? sqrt : -sqrt));
        double t1 = q / a;
        double t2 = q == 0D ? 0D : c / q;
        // Squaring the circle equation also produces roots with negative radii.
        // Only physical circles contribute; where two exist, the later circle wins.
        bool valid1 = startRadius + t1 * dr >= 0D;
        bool valid2 = startRadius + t2 * dr >= 0D;
        if (!valid1 && !valid2) return MissingCircleRatio;
        double ratio = valid1 ? (valid2 ? (pointEnd ? Math.Min(t1, t2) : Math.Max(t1, t2)) : t1) : t2;
        return pointEnd ? 1D - ratio : ratio;
    }

}
