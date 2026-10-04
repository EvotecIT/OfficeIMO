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

    internal double SampleRatio(double x, double y) => Math.Max(0D, Math.Min(1D, SampleUnboundedRatio(x, y)));

    internal double SampleUnboundedRatio(double x, double y) {
        var point = _inverseCoordinates.TransformPoint(new OfficePoint(x, y));
        x = point.X; y = point.Y;
        double endRadiusX = Math.Max(Math.Max(StartRadiusX, EndRadiusX), 0.0000001D);
        double endRadiusY = Math.Max(Math.Max(StartRadiusY, EndRadiusY), 0.0000001D);
        double normalizedX = (x - EndX) / endRadiusX;
        double normalizedY = (y - EndY) / endRadiusY;
        double startX = (StartX - EndX) / endRadiusX;
        double startY = (StartY - EndY) / endRadiusY;
        double startRadius = StartRadiusX / endRadiusX;
        double vx = normalizedX - startX;
        double vy = normalizedY - startY;
        double dx = -startX;
        double dy = -startY;
        double dr = (EndRadiusX - StartRadiusX) / endRadiusX;
        double a = (dx * dx) + (dy * dy) - (dr * dr);
        double b = -2D * ((vx * dx) + (vy * dy) + (startRadius * dr));
        double c = (vx * vx) + (vy * vy) - (startRadius * startRadius);
        if (Math.Abs(a) < 0.0000001D) {
            if (Math.Abs(b) < 0.0000001D) {
                return 0D;
            }

            double linearRatio = -c / b;
            return startRadius + linearRatio * dr >= 0D ? linearRatio : 0D;
        }

        double discriminant = (b * b) - (4D * a * c);
        if (discriminant < 0D) {
            return 0D;
        }

        double sqrt = Math.Sqrt(discriminant);
        double t1 = (-b - sqrt) / (2D * a);
        double t2 = (-b + sqrt) / (2D * a);
        // Squaring the circle equation also produces roots with negative radii.
        // Only physical circles contribute; where two exist, the later circle wins.
        bool valid1 = startRadius + t1 * dr >= 0D;
        bool valid2 = startRadius + t2 * dr >= 0D;
        return valid1 ? (valid2 ? Math.Max(t1, t2) : t1) : (valid2 ? t2 : 0D);
    }

}
