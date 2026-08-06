using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free radial gradient intent in normalized local coordinates.
/// Coordinates use a top-left origin where 0,0 is the shape's top-left corner and 1,1 is its bottom-right corner.
/// Ellipse centers may sit outside that box to preserve authored gradients whose focal point is off-canvas.
/// Independent horizontal and vertical radii allow axis-aligned ellipses while the original circle constructors remain source-compatible.
/// </summary>
public sealed class OfficeRadialGradient {
    /// <summary>Start circle center X coordinate in shape-local units.</summary>
    public double StartX { get; }

    /// <summary>Start circle center Y coordinate in shape-local units.</summary>
    public double StartY { get; }

    /// <summary>Normalized start circle radius, or the horizontal radius for an elliptical definition.</summary>
    public double StartRadius { get; }

    /// <summary>Normalized horizontal start ellipse radius.</summary>
    public double StartRadiusX { get; }

    /// <summary>Normalized vertical start ellipse radius.</summary>
    public double StartRadiusY { get; }

    /// <summary>End circle center X coordinate in shape-local units.</summary>
    public double EndX { get; }

    /// <summary>End circle center Y coordinate in shape-local units.</summary>
    public double EndY { get; }

    /// <summary>Normalized end circle radius, or the horizontal radius for an elliptical definition.</summary>
    public double EndRadius { get; }

    /// <summary>Normalized horizontal end ellipse radius.</summary>
    public double EndRadiusX { get; }

    /// <summary>Normalized vertical end ellipse radius.</summary>
    public double EndRadiusY { get; }

    /// <summary>Gradient stops in offset order.</summary>
    public IReadOnlyList<OfficeGradientStop> Stops { get; }

    /// <summary>Creates a radial gradient between two circles.</summary>
    public OfficeRadialGradient(double startX, double startY, double startRadius, double endX, double endY, double endRadius, OfficeGradientStop start, OfficeGradientStop end) {
        ValidateCoordinates(startX, startY, startRadius, endX, endY, endRadius);
        StartX = startX;
        StartY = startY;
        StartRadius = startRadius;
        StartRadiusX = startRadius;
        StartRadiusY = startRadius;
        EndX = endX;
        EndY = endY;
        EndRadius = endRadius;
        EndRadiusX = endRadius;
        EndRadiusY = endRadius;
        Stops = ValidateStops(new[] { start, end });
    }

    /// <summary>Creates a radial gradient between two circles with two or more ordered stops.</summary>
    public OfficeRadialGradient(double startX, double startY, double startRadius, double endX, double endY, double endRadius, IReadOnlyList<OfficeGradientStop> stops) {
        ValidateCoordinates(startX, startY, startRadius, endX, endY, endRadius);
        StartX = startX;
        StartY = startY;
        StartRadius = startRadius;
        StartRadiusX = startRadius;
        StartRadiusY = startRadius;
        EndX = endX;
        EndY = endY;
        EndRadius = endRadius;
        EndRadiusX = endRadius;
        EndRadiusY = endRadius;
        Stops = ValidateStops(stops);
    }

    /// <summary>Creates a radial gradient between two axis-aligned ellipses with the same aspect ratio.</summary>
    public OfficeRadialGradient(
        double startX,
        double startY,
        double startRadiusX,
        double startRadiusY,
        double endX,
        double endY,
        double endRadiusX,
        double endRadiusY,
        IReadOnlyList<OfficeGradientStop> stops) {
        ValidateCoordinates(startX, startY, startRadiusX, startRadiusY, endX, endY, endRadiusX, endRadiusY);
        StartX = startX;
        StartY = startY;
        StartRadius = startRadiusX;
        StartRadiusX = startRadiusX;
        StartRadiusY = startRadiusY;
        EndX = endX;
        EndY = endY;
        EndRadius = endRadiusX;
        EndRadiusX = endRadiusX;
        EndRadiusY = endRadiusY;
        Stops = ValidateStops(stops);
    }

    /// <summary>Creates a centered radial gradient from the center outward.</summary>
    public static OfficeRadialGradient Centered(OfficeColor startColor, OfficeColor endColor) =>
        new OfficeRadialGradient(0.5D, 0.5D, 0D, 0.5D, 0.5D, 0.5D, new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor));

    /// <summary>Creates a detached copy.</summary>
    public OfficeRadialGradient Clone() => new OfficeRadialGradient(StartX, StartY, StartRadiusX, StartRadiusY, EndX, EndY, EndRadiusX, EndRadiusY, Stops);

    private static void ValidateCoordinates(double startX, double startY, double startRadius, double endX, double endY, double endRadius) {
        ValidateFiniteCoordinate(startX, nameof(startX));
        ValidateFiniteCoordinate(startY, nameof(startY));
        ValidateFiniteCoordinate(endX, nameof(endX));
        ValidateFiniteCoordinate(endY, nameof(endY));
        ValidateRadius(startRadius, nameof(startRadius));
        ValidateRadius(endRadius, nameof(endRadius));
        if (startX.Equals(endX) && startY.Equals(endY) && startRadius.Equals(endRadius)) {
            throw new ArgumentException("Radial gradient start and end circles must be different.");
        }
    }

    private static void ValidateCoordinates(
        double startX,
        double startY,
        double startRadiusX,
        double startRadiusY,
        double endX,
        double endY,
        double endRadiusX,
        double endRadiusY) {
        ValidateFiniteCoordinate(startX, nameof(startX));
        ValidateFiniteCoordinate(startY, nameof(startY));
        ValidateFiniteCoordinate(endX, nameof(endX));
        ValidateFiniteCoordinate(endY, nameof(endY));
        ValidateRadius(startRadiusX, nameof(startRadiusX));
        ValidateRadius(startRadiusY, nameof(startRadiusY));
        ValidateRadius(endRadiusX, nameof(endRadiusX));
        ValidateRadius(endRadiusY, nameof(endRadiusY));
        if (endRadiusX <= 0D) throw new ArgumentOutOfRangeException(nameof(endRadiusX), "Radial gradient end ellipse radii must be positive.");
        if (endRadiusY <= 0D) throw new ArgumentOutOfRangeException(nameof(endRadiusY), "Radial gradient end ellipse radii must be positive.");

        if ((startRadiusX > 0D || startRadiusY > 0D)
            && Math.Abs((startRadiusX / endRadiusX) - (startRadiusY / endRadiusY)) > 0.0000001D) {
            throw new ArgumentException("Radial gradient start and end ellipses must use the same aspect ratio.");
        }

        if (startX.Equals(endX)
            && startY.Equals(endY)
            && startRadiusX.Equals(endRadiusX)
            && startRadiusY.Equals(endRadiusY)) {
            throw new ArgumentException("Radial gradient start and end ellipses must be different.");
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
            if (stop.Offset < previous) {
                throw new ArgumentException("Radial gradient stops must be in non-decreasing offset order.", nameof(stops));
            }

            copy.Add(stop);
            previous = stop.Offset;
        }

        return new ReadOnlyCollection<OfficeGradientStop>(copy);
    }

    private static void ValidateFiniteCoordinate(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(name, "Radial gradient coordinates must be finite values.");
        }
    }

    private static void ValidateRadius(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(name, "Radial gradient radii must be finite non-negative values.");
        }
    }
}
