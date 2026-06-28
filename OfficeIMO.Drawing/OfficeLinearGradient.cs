using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free linear gradient intent in normalized local coordinates.
/// Coordinates use a top-left origin where 0,0 is the shape's top-left corner and 1,1 is its bottom-right corner.
/// </summary>
public sealed class OfficeLinearGradient {
    /// <summary>Normalized start X coordinate.</summary>
    public double StartX { get; }

    /// <summary>Normalized start Y coordinate.</summary>
    public double StartY { get; }

    /// <summary>Normalized end X coordinate.</summary>
    public double EndX { get; }

    /// <summary>Normalized end Y coordinate.</summary>
    public double EndY { get; }

    /// <summary>Gradient stops in offset order.</summary>
    public IReadOnlyList<OfficeGradientStop> Stops { get; }

    /// <summary>Creates a two-stop linear gradient.</summary>
    public OfficeLinearGradient(double startX, double startY, double endX, double endY, OfficeGradientStop start, OfficeGradientStop end) {
        ValidateNormalized(startX, nameof(startX));
        ValidateNormalized(startY, nameof(startY));
        ValidateNormalized(endX, nameof(endX));
        ValidateNormalized(endY, nameof(endY));

        if (startX.Equals(endX) && startY.Equals(endY)) {
            throw new ArgumentException("Linear gradient start and end points must be different.");
        }

        if (!start.Offset.Equals(0D)) {
            throw new ArgumentException("The first linear gradient stop must use offset 0.", nameof(start));
        }

        if (!end.Offset.Equals(1D)) {
            throw new ArgumentException("The second linear gradient stop must use offset 1.", nameof(end));
        }

        StartX = startX;
        StartY = startY;
        EndX = endX;
        EndY = endY;
        Stops = new ReadOnlyCollection<OfficeGradientStop>(new List<OfficeGradientStop> { start, end });
    }

    /// <summary>Creates a horizontal left-to-right gradient.</summary>
    public static OfficeLinearGradient Horizontal(OfficeColor startColor, OfficeColor endColor) =>
        new OfficeLinearGradient(0, 0.5, 1, 0.5, new OfficeGradientStop(0, startColor), new OfficeGradientStop(1, endColor));

    /// <summary>Creates a vertical top-to-bottom gradient.</summary>
    public static OfficeLinearGradient Vertical(OfficeColor startColor, OfficeColor endColor) =>
        new OfficeLinearGradient(0.5, 0, 0.5, 1, new OfficeGradientStop(0, startColor), new OfficeGradientStop(1, endColor));

    /// <summary>Creates a diagonal top-left to bottom-right gradient.</summary>
    public static OfficeLinearGradient DiagonalDown(OfficeColor startColor, OfficeColor endColor) =>
        new OfficeLinearGradient(0, 0, 1, 1, new OfficeGradientStop(0, startColor), new OfficeGradientStop(1, endColor));

    /// <summary>
    /// Creates a two-stop gradient projected through the normalized drawing rectangle at the supplied angle.
    /// </summary>
    /// <param name="startColor">Color at offset 0.</param>
    /// <param name="endColor">Color at offset 1.</param>
    /// <param name="degrees">Clockwise angle in degrees where 0 is left-to-right in local coordinates.</param>
    /// <returns>A gradient with endpoints clamped to the normalized drawing rectangle.</returns>
    public static OfficeLinearGradient FromAngle(OfficeColor startColor, OfficeColor endColor, double degrees) {
        if (double.IsNaN(degrees) || double.IsInfinity(degrees)) {
            throw new ArgumentOutOfRangeException(nameof(degrees), "Linear gradient angle must be finite.");
        }

        double radians = OfficeGeometry.DegreesToRadians(NormalizeDegrees(degrees));
        double dx = Math.Cos(radians);
        double dy = Math.Sin(radians);
        double divisor = Math.Max(Math.Abs(dx), Math.Abs(dy));
        if (divisor <= double.Epsilon) {
            return Horizontal(startColor, endColor);
        }

        double half = 0.5D / divisor;
        return new OfficeLinearGradient(
            ClampUnit(0.5D - (dx * half)),
            ClampUnit(0.5D - (dy * half)),
            ClampUnit(0.5D + (dx * half)),
            ClampUnit(0.5D + (dy * half)),
            new OfficeGradientStop(0D, startColor),
            new OfficeGradientStop(1D, endColor));
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeLinearGradient Clone() => new OfficeLinearGradient(StartX, StartY, EndX, EndY, Stops[0], Stops[1]);

    private static double NormalizeDegrees(double degree) {
        double normalized = degree % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }

    private static double ClampUnit(double value) => value < 0D ? 0D : value > 1D ? 1D : value;

    private static void ValidateNormalized(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(paramName, "Linear gradient coordinates must be finite values between 0 and 1.");
        }
    }
}
