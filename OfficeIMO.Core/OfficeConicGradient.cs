using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Backend-neutral conic gradient intent that can be expanded into a bounded vector drawing.
/// Stop offsets cover one clockwise turn, and zero degrees points upward.
/// </summary>
public sealed class OfficeConicGradient {
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

    /// <summary>
    /// Expands the gradient into clipped solid-color vector wedges shared by raster, SVG, and PDF exporters.
    /// Authored stop boundaries are retained in addition to the uniform quality segments.
    /// </summary>
    public OfficeDrawing CreateDrawing(double width, double height, int qualitySegments = 360) {
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        if (qualitySegments < 12 || qualitySegments > 4096) {
            throw new ArgumentOutOfRangeException(nameof(qualitySegments), "Conic gradient quality must be between 12 and 4096 segments.");
        }

        var boundaries = new SortedSet<double>();
        for (int index = 0; index <= qualitySegments; index++) boundaries.Add((double)index / qualitySegments);
        foreach (OfficeGradientStop stop in Stops) boundaries.Add(stop.Offset);
        var ordered = new List<double>(boundaries);
        double centerX = CenterX * width;
        double centerY = CenterY * height;
        double radius = 2D * Math.Sqrt((width * width) + (height * height));
        var content = new OfficeDrawing(width, height);
        const double overlap = 0.000001D;
        for (int index = 1; index < ordered.Count; index++) {
            double start = ordered[index - 1];
            double end = ordered[index];
            if (end <= start) continue;
            double startRadians = ToRadians(StartAngle + (start * 360D)) - (Math.PI / 2D) - overlap;
            double endRadians = ToRadians(StartAngle + (end * 360D)) - (Math.PI / 2D) + overlap;
            double firstX = centerX + (Math.Cos(startRadians) * radius);
            double firstY = centerY + (Math.Sin(startRadians) * radius);
            double secondX = centerX + (Math.Cos(endRadians) * radius);
            double secondY = centerY + (Math.Sin(endRadians) * radius);
            OfficeShape wedge = OfficeShape.Path(
                width,
                height,
                OfficePathCommand.MoveTo(centerX, centerY),
                OfficePathCommand.LineTo(firstX, firstY),
                OfficePathCommand.LineTo(secondX, secondY),
                OfficePathCommand.Close());
            wedge.FillColor = Sample((start + end) / 2D);
            wedge.StrokeWidth = 0D;
            content.AddShapeForClippedRendering(wedge, 0D, 0D);
        }

        var drawing = new OfficeDrawing(width, height);
        drawing.AddClippedDrawing(content, 0D, 0D, OfficeClipPath.Rectangle(width, height));
        return drawing;
    }

    /// <summary>Creates a detached copy.</summary>
    public OfficeConicGradient Clone() => new OfficeConicGradient(CenterX, CenterY, StartAngle, Stops);

    private OfficeColor Sample(double offset) {
        if (offset <= Stops[0].Offset) return Stops[0].Color;
        for (int index = 1; index < Stops.Count; index++) {
            OfficeGradientStop current = Stops[index];
            if (offset > current.Offset) continue;
            OfficeGradientStop previous = Stops[index - 1];
            if (current.Offset <= previous.Offset) return current.Color;
            double ratio = (offset - previous.Offset) / (current.Offset - previous.Offset);
            return OfficeColor.FromRgba(
                Interpolate(previous.Color.R, current.Color.R, ratio),
                Interpolate(previous.Color.G, current.Color.G, ratio),
                Interpolate(previous.Color.B, current.Color.B, ratio),
                Interpolate(previous.Color.A, current.Color.A, ratio));
        }
        return Stops[Stops.Count - 1].Color;
    }

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

    private static double ToRadians(double degrees) => degrees * Math.PI / 180D;

    private static void ValidateFinite(double value, string parameter) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(parameter, "Conic gradient values must be finite.");
    }

    private static void ValidatePositive(double value, string parameter) {
        ValidateFinite(value, parameter);
        if (value <= 0D) throw new ArgumentOutOfRangeException(parameter, "Conic gradient drawing dimensions must be positive.");
    }
}
