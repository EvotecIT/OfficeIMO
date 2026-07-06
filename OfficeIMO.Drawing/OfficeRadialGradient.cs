using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free radial gradient intent in normalized local coordinates.
/// Coordinates use a top-left origin where 0,0 is the shape's top-left corner and 1,1 is its bottom-right corner.
/// </summary>
public sealed class OfficeRadialGradient {
    /// <summary>Normalized start circle center X coordinate.</summary>
    public double StartX { get; }

    /// <summary>Normalized start circle center Y coordinate.</summary>
    public double StartY { get; }

    /// <summary>Normalized start circle radius.</summary>
    public double StartRadius { get; }

    /// <summary>Normalized end circle center X coordinate.</summary>
    public double EndX { get; }

    /// <summary>Normalized end circle center Y coordinate.</summary>
    public double EndY { get; }

    /// <summary>Normalized end circle radius.</summary>
    public double EndRadius { get; }

    /// <summary>Gradient stops in offset order.</summary>
    public IReadOnlyList<OfficeGradientStop> Stops { get; }

    /// <summary>Creates a radial gradient between two circles.</summary>
    public OfficeRadialGradient(double startX, double startY, double startRadius, double endX, double endY, double endRadius, OfficeGradientStop start, OfficeGradientStop end) {
        ValidateCoordinates(startX, startY, startRadius, endX, endY, endRadius);
        StartX = startX;
        StartY = startY;
        StartRadius = startRadius;
        EndX = endX;
        EndY = endY;
        EndRadius = endRadius;
        Stops = ValidateStops(new[] { start, end });
    }

    /// <summary>Creates a radial gradient between two circles with two or more ordered stops.</summary>
    public OfficeRadialGradient(double startX, double startY, double startRadius, double endX, double endY, double endRadius, IReadOnlyList<OfficeGradientStop> stops) {
        ValidateCoordinates(startX, startY, startRadius, endX, endY, endRadius);
        StartX = startX;
        StartY = startY;
        StartRadius = startRadius;
        EndX = endX;
        EndY = endY;
        EndRadius = endRadius;
        Stops = ValidateStops(stops);
    }

    /// <summary>Creates a centered radial gradient from the center outward.</summary>
    public static OfficeRadialGradient Centered(OfficeColor startColor, OfficeColor endColor) =>
        new OfficeRadialGradient(0.5D, 0.5D, 0D, 0.5D, 0.5D, 0.5D, new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor));

    /// <summary>Creates a detached copy.</summary>
    public OfficeRadialGradient Clone() => new OfficeRadialGradient(StartX, StartY, StartRadius, EndX, EndY, EndRadius, Stops);

    private static void ValidateCoordinates(double startX, double startY, double startRadius, double endX, double endY, double endRadius) {
        ValidateNormalized(startX, nameof(startX));
        ValidateNormalized(startY, nameof(startY));
        ValidateNormalized(endX, nameof(endX));
        ValidateNormalized(endY, nameof(endY));
        ValidateRadius(startRadius, nameof(startRadius));
        ValidateRadius(endRadius, nameof(endRadius));
        if (startX.Equals(endX) && startY.Equals(endY) && startRadius.Equals(endRadius)) {
            throw new ArgumentException("Radial gradient start and end circles must be different.");
        }
    }

    private static IReadOnlyList<OfficeGradientStop> ValidateStops(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null || stops.Count < 2) {
            throw new ArgumentException("A radial gradient needs at least two stops.", nameof(stops));
        }

        if (!stops[0].Offset.Equals(0D)) {
            throw new ArgumentException("The first radial gradient stop must use offset 0.", nameof(stops));
        }

        if (!stops[stops.Count - 1].Offset.Equals(1D)) {
            throw new ArgumentException("The last radial gradient stop must use offset 1.", nameof(stops));
        }

        var copy = new List<OfficeGradientStop>(stops.Count);
        double previous = -1D;
        for (int i = 0; i < stops.Count; i++) {
            OfficeGradientStop stop = stops[i];
            if (stop.Offset <= previous) {
                throw new ArgumentException("Radial gradient stops must be in strictly increasing offset order.", nameof(stops));
            }

            copy.Add(stop);
            previous = stop.Offset;
        }

        return new ReadOnlyCollection<OfficeGradientStop>(copy);
    }

    private static void ValidateNormalized(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(name, "Radial gradient coordinates must be finite values between 0 and 1.");
        }
    }

    private static void ValidateRadius(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(name, "Radial gradient radii must be finite non-negative values.");
        }
    }
}
