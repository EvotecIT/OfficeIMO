using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Backend-neutral conic gradient intent that can be expanded into a bounded vector drawing.
/// Stop offsets cover one clockwise turn, and zero degrees points upward.
/// </summary>
public sealed partial class OfficeConicGradient {
    /// <summary>Normalized horizontal center coordinate.</summary>
    public double CenterX { get; }

    /// <summary>Normalized vertical center coordinate.</summary>
    public double CenterY { get; }

    /// <summary>Clockwise start angle in degrees where zero points upward.</summary>
    public double StartAngle { get; }

    /// <summary>Gradient stops in non-decreasing offset order from zero to one.</summary>
    public IReadOnlyList<OfficeGradientStop> Stops { get; }

    /// <summary>Creates a conic gradient in normalized local coordinates.</summary>
    public OfficeConicGradient(
        double centerX,
        double centerY,
        double startAngle,
        IReadOnlyList<OfficeGradientStop> stops) {
        ValidateFinite(centerX, nameof(centerX));
        ValidateFinite(centerY, nameof(centerY));
        ValidateFinite(startAngle, nameof(startAngle));
        CenterX = centerX;
        CenterY = centerY;
        StartAngle = NormalizeDegrees(startAngle);
        Stops = ValidateStops(stops);
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeConicGradient Clone() => new OfficeConicGradient(CenterX, CenterY, StartAngle, Stops);

    internal OfficeColor Sample(double offset) {
        if (offset <= Stops[0].Offset) return Stops[0].Color;
        for (int index = 1; index < Stops.Count; index++) {
            OfficeGradientStop current = Stops[index];
            if (offset > current.Offset) continue;
            OfficeGradientStop previous = Stops[index - 1];
            if (current.Offset <= previous.Offset) return current.Color;
            double ratio = (offset - previous.Offset) / (current.Offset - previous.Offset);
            return InterpolateColor(previous.Color, current.Color, ratio);
        }
        return Stops[Stops.Count - 1].Color;
    }

    private static OfficeColor InterpolateColor(OfficeColor first, OfficeColor second, double ratio) {
        double firstAlpha = first.A / 255D;
        double secondAlpha = second.A / 255D;
        double alpha = firstAlpha + ((secondAlpha - firstAlpha) * ratio);
        byte outputAlpha = Interpolate(first.A, second.A, ratio);
        if (alpha <= 0.0000001D || outputAlpha == 0) return OfficeColor.FromRgba(0, 0, 0, 0);

        return OfficeColor.FromRgba(
            Unpremultiply(first.R * firstAlpha + ((second.R * secondAlpha - first.R * firstAlpha) * ratio), alpha),
            Unpremultiply(first.G * firstAlpha + ((second.G * secondAlpha - first.G * firstAlpha) * ratio), alpha),
            Unpremultiply(first.B * firstAlpha + ((second.B * secondAlpha - first.B * firstAlpha) * ratio), alpha),
            outputAlpha);
    }

    private static byte Unpremultiply(double value, double alpha) =>
        (byte)Math.Max(0D, Math.Min(255D, Math.Round(value / alpha, MidpointRounding.AwayFromZero)));

    private static byte Interpolate(byte first, byte second, double ratio) =>
        (byte)Math.Round(first + ((second - first) * ratio), MidpointRounding.AwayFromZero);

    private static IReadOnlyList<OfficeGradientStop> ValidateStops(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null || stops.Count < 2) throw new ArgumentException("A conic gradient needs at least two stops.", nameof(stops));
        if (!stops[0].Offset.Equals(0D) || !stops[stops.Count - 1].Offset.Equals(1D)) {
            throw new ArgumentException("Conic gradient stops must cover offsets zero through one.", nameof(stops));
        }
        var copy = new List<OfficeGradientStop>(stops.Count);
        double previous = -1D;
        foreach (OfficeGradientStop stop in stops) {
            if (stop.Offset < previous) throw new ArgumentException("Conic gradient stops must be ordered.", nameof(stops));
            copy.Add(stop);
            previous = stop.Offset;
        }
        return new ReadOnlyCollection<OfficeGradientStop>(copy);
    }

    private static double NormalizeDegrees(double value) {
        double normalized = value % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }

    private static void ValidateFinite(double value, string parameter) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(parameter, "Conic gradient values must be finite.");
    }

    private static void ValidatePositive(double value, string parameter) {
        ValidateFinite(value, parameter);
        if (value <= 0D) throw new ArgumentOutOfRangeException(parameter, "Conic gradient drawing dimensions must be positive.");
    }
}
