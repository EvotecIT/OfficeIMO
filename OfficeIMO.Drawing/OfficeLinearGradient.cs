using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free linear gradient intent in normalized local coordinates.
/// Coordinates use a top-left origin where 0,0 is the shape's top-left corner and 1,1 is its bottom-right corner.
/// Importers may preserve finite user-space vectors that cross those bounds without weakening public constructor validation.
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
    public OfficeLinearGradient(double startX, double startY, double endX, double endY, OfficeGradientStop start, OfficeGradientStop end)
        : this(startX, startY, endX, endY, new[] { start, end }, allowOutsideUnit: false) {
    }

    /// <summary>Creates a linear gradient with two or more ordered stops.</summary>
    public OfficeLinearGradient(double startX, double startY, double endX, double endY, IReadOnlyList<OfficeGradientStop> stops)
        : this(startX, startY, endX, endY, stops, allowOutsideUnit: false) {
    }

    private OfficeLinearGradient(
        double startX,
        double startY,
        double endX,
        double endY,
        IReadOnlyList<OfficeGradientStop> stops,
        bool allowOutsideUnit) {
        ValidateCoordinates(startX, startY, endX, endY, allowOutsideUnit);
        StartX = startX;
        StartY = startY;
        EndX = endX;
        EndY = endY;
        Stops = ValidateStops(stops);
    }

    private static void ValidateCoordinates(double startX, double startY, double endX, double endY, bool allowOutsideUnit) {
        ValidateCoordinate(startX, nameof(startX), allowOutsideUnit);
        ValidateCoordinate(startY, nameof(startY), allowOutsideUnit);
        ValidateCoordinate(endX, nameof(endX), allowOutsideUnit);
        ValidateCoordinate(endY, nameof(endY), allowOutsideUnit);

        if (startX.Equals(endX) && startY.Equals(endY)) {
            throw new ArgumentException("Linear gradient start and end points must be different.");
        }
    }

    internal static OfficeLinearGradient CreateImported(
        double startX,
        double startY,
        double endX,
        double endY,
        IReadOnlyList<OfficeGradientStop> stops) =>
        new OfficeLinearGradient(startX, startY, endX, endY, stops, allowOutsideUnit: true);

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
        return FromAngle(
            new[] {
                new OfficeGradientStop(0D, startColor),
                new OfficeGradientStop(1D, endColor)
            },
            degrees);
    }

    /// <summary>
    /// Creates a gradient with ordered stops projected through the normalized drawing rectangle at the supplied angle.
    /// </summary>
    /// <param name="stops">Gradient stops. The first stop must be offset 0 and the last stop must be offset 1.</param>
    /// <param name="degrees">Clockwise angle in degrees where 0 is left-to-right in local coordinates.</param>
    /// <returns>A gradient with endpoints clamped to the normalized drawing rectangle.</returns>
    public static OfficeLinearGradient FromAngle(IReadOnlyList<OfficeGradientStop> stops, double degrees) {
        if (double.IsNaN(degrees) || double.IsInfinity(degrees)) {
            throw new ArgumentOutOfRangeException(nameof(degrees), "Linear gradient angle must be finite.");
        }

        double radians = OfficeGeometry.DegreesToRadians(NormalizeDegrees(degrees));
        double dx = Math.Cos(radians);
        double dy = Math.Sin(radians);
        double divisor = Math.Max(Math.Abs(dx), Math.Abs(dy));
        if (divisor <= double.Epsilon) {
            return new OfficeLinearGradient(0, 0.5, 1, 0.5, stops);
        }

        double half = 0.5D / divisor;
        return new OfficeLinearGradient(
            ClampUnit(0.5D - (dx * half)),
            ClampUnit(0.5D - (dy * half)),
            ClampUnit(0.5D + (dx * half)),
            ClampUnit(0.5D + (dy * half)),
            stops);
    }

    /// <summary>
    /// Creates a local-coordinate gradient whose vector follows the supplied angle after an affine shape transform.
    /// </summary>
    /// <param name="stops">Gradient stops. The first stop must use offset 0 and the last stop offset 1.</param>
    /// <param name="destinationDegrees">Clockwise gradient angle in destination coordinates.</param>
    /// <param name="localWidth">Width of the shape's local coordinate box.</param>
    /// <param name="localHeight">Height of the shape's local coordinate box.</param>
    /// <param name="localToDestination">Affine transform applied to the shape after its local gradient is evaluated.</param>
    /// <returns>A local normalized gradient that preserves the requested destination-space direction.</returns>
    public static OfficeLinearGradient FromTransformedAngle(
        IReadOnlyList<OfficeGradientStop> stops,
        double destinationDegrees,
        double localWidth,
        double localHeight,
        OfficeTransform localToDestination) {
        if (double.IsNaN(destinationDegrees) || double.IsInfinity(destinationDegrees)) {
            throw new ArgumentOutOfRangeException(nameof(destinationDegrees),
                "Linear gradient angle must be finite.");
        }
        ValidatePositiveDimension(localWidth, nameof(localWidth));
        ValidatePositiveDimension(localHeight, nameof(localHeight));
        if (!localToDestination.TryInvert(out OfficeTransform destinationToLocal)) {
            throw new ArgumentException("The local-to-destination transform must be invertible.",
                nameof(localToDestination));
        }

        double radians = OfficeGeometry.DegreesToRadians(
            NormalizeDegrees(destinationDegrees));
        OfficePoint localOrigin = destinationToLocal.TransformPoint(default);
        OfficePoint localDirection = destinationToLocal.TransformPoint(new OfficePoint(
            Math.Cos(radians), Math.Sin(radians)));
        double normalizedX = (localDirection.X - localOrigin.X) / localWidth;
        double normalizedY = (localDirection.Y - localOrigin.Y) / localHeight;
        double localDegrees = Math.Atan2(normalizedY, normalizedX) * 180D / Math.PI;
        return FromAngle(stops, localDegrees);
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeLinearGradient Clone() => new OfficeLinearGradient(StartX, StartY, EndX, EndY, Stops, allowOutsideUnit: true);

    private static IReadOnlyList<OfficeGradientStop> ValidateStops(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null || stops.Count < 2) {
            throw new ArgumentException("A linear gradient needs at least two stops.", nameof(stops));
        }

        if (!stops[0].Offset.Equals(0D)) {
            throw new ArgumentException("The first linear gradient stop must use offset 0.", nameof(stops));
        }

        if (!stops[stops.Count - 1].Offset.Equals(1D)) {
            throw new ArgumentException("The last linear gradient stop must use offset 1.", nameof(stops));
        }

        var copy = new List<OfficeGradientStop>(stops.Count);
        double previous = -1D;
        for (int i = 0; i < stops.Count; i++) {
            OfficeGradientStop stop = stops[i];
            if (stop.Offset < previous) {
                throw new ArgumentException("Linear gradient stops must be in non-decreasing offset order.", nameof(stops));
            }

            copy.Add(stop);
            previous = stop.Offset;
        }

        return new ReadOnlyCollection<OfficeGradientStop>(copy);
    }

    private static double NormalizeDegrees(double degree) {
        double normalized = degree % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }

    private static double ClampUnit(double value) => value < 0D ? 0D : value > 1D ? 1D : value;

    private static void ValidatePositiveDimension(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName,
                "Linear gradient dimensions must be finite positive values.");
        }
    }

    private static void ValidateCoordinate(double value, string paramName, bool allowOutsideUnit) {
        if (double.IsNaN(value) || double.IsInfinity(value) || !allowOutsideUnit && (value < 0D || value > 1D)) {
            throw new ArgumentOutOfRangeException(paramName, allowOutsideUnit
                ? "Linear gradient coordinates must be finite values."
                : "Linear gradient coordinates must be finite values between 0 and 1.");
        }
    }
}
