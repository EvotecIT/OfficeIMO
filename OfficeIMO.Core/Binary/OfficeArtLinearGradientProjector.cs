using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

// The native OfficeArt scalar field belongs with the shared OfficeArt codec.
// Application importers own palette resolution and operation-level loss reports.
internal static class OfficeArtLinearGradientProjector {
    internal static bool TryProject(OfficeArtShapeStyle style, double width, double height,
        Func<OfficeArtColorReference, OfficeColor?> resolve, out OfficeArtLinearGradientFill? fill, out string? failure) {
        fill = null; failure = null;
        if (style.FillType is not (4U or 7U)) { failure = "This native fill type is not linear."; return false; }
        if (style.FillUsesCustomRectangle == true || style.FillAlignedWithShape == false) {
            failure = "The native gradient uses a custom fill rectangle or view-origin anchor."; return false;
        }
        int focus = style.FillFocusPercent.GetValueOrDefault();
        if (focus < -100 || focus > 100 || style.IsFillGradientStopTableTruncated
            || InvalidOpacity(style, 0x0182) || InvalidOpacity(style, 0x0184)) {
            failure = "The native gradient focus, opacity, or color-stop table is invalid."; return false;
        }
        if (width <= 0 || height <= 0 || double.IsInfinity(width) || double.IsInfinity(height)
            || double.IsNaN(width) || double.IsNaN(height)) {
            failure = "The gradient frame has invalid dimensions."; return false;
        }
        var stops = new List<OfficeGradientStop>();
        double foregroundOpacity = style.FillOpacity ?? 1, backgroundOpacity = style.FillBackOpacity ?? 1;
        double opacity = style.FillGradientStops.Count > 0 ? foregroundOpacity : Math.Max(foregroundOpacity, backgroundOpacity);
        bool opacityApproximated = false;
        if (style.FillGradientStops.Count > 0) {
            foreach (OfficeArtGradientStop stop in style.FillGradientStops)
                stops.Add(new OfficeGradientStop(stop.Position, Opacity(resolve(stop.Color) ?? OfficeColor.White,
                    foregroundOpacity, opacity, ref opacityApproximated)));
            // Missing endpoint entries use the nearest declared color, as a
            // padded gradient does. Explicit duplicate positions stay ordered.
            if (stops[0].Offset > 0) stops.Insert(0, new OfficeGradientStop(0, stops[0].Color));
            if (stops[stops.Count - 1].Offset < 1) stops.Add(new OfficeGradientStop(1, stops[stops.Count - 1].Color));
        } else {
            stops.Add(new OfficeGradientStop(0, Opacity(Resolve(style.FillColor, resolve), foregroundOpacity, opacity, ref opacityApproximated)));
            stops.Add(new OfficeGradientStop(1, Opacity(Resolve(style.FillBackColor, resolve), backgroundOpacity, opacity, ref opacityApproximated)));
        }
        stops = ApplyFocus(stops, focus);
        double radians = (style.FillAngleDegrees.GetValueOrDefault() % 360) * Math.PI / 180;
        double nx = -Math.Sin(radians), ny = -Math.Cos(radians);
        if (Math.Abs(nx) < 1E-14) nx = 0;
        if (Math.Abs(ny) < 1E-14) ny = 0;
        if (style.FillType == 4) {
            // A physical angle is a normal to the color field. Shape-local
            // normalized coordinates scale that normal by the frame dimensions.
            double scale = Math.Max(width, height);
            nx *= width / scale; ny *= height / scale;
        }
        double span = Math.Abs(nx) + Math.Abs(ny);
        if (span == 0) { failure = "The gradient vector is not representable."; return false; }
        nx /= span; ny /= span;
        double squared = nx * nx + ny * ny;
        double dx = nx / squared, dy = ny / squared;
        OfficeLinearGradient gradient = OfficeLinearGradient.CreateImported(0.5 - dx / 2, 0.5 - dy / 2, 0.5 + dx / 2, 0.5 + dy / 2, stops)
            .WithSeparateAlphaInterpolation(true);
        fill = new OfficeArtLinearGradientFill(gradient, opacity, opacityApproximated);
        return true;
    }

    private static OfficeColor Resolve(OfficeArtColorReference? reference, Func<OfficeArtColorReference, OfficeColor?> resolve) =>
        reference.HasValue ? resolve(reference.Value) ?? OfficeColor.White : OfficeColor.White;

    private static OfficeColor Opacity(OfficeColor color, double opacity, double commonOpacity, ref bool approximated) {
        double alpha = commonOpacity == 0 ? 0 : color.A * opacity / commonOpacity;
        byte rounded = (byte)Math.Round(alpha, MidpointRounding.AwayFromZero);
        approximated |= Math.Abs(rounded - alpha) > 1E-9;
        return OfficeColor.FromRgba(color.R, color.G, color.B, rounded);
    }

    private static bool InvalidOpacity(OfficeArtShapeStyle style, ushort id) => style.Properties.LastOrDefault(property =>
        property.PropertyId == id && !property.IsComplex)?.Value > 65536;

    private static List<OfficeGradientStop> ApplyFocus(List<OfficeGradientStop> stops, int focus) {
        if (focus == 0) return stops;
        if (focus is 100 or -100) return stops.AsEnumerable().Reverse()
            .Select(stop => new OfficeGradientStop(1 - stop.Offset, stop.Color)).ToList();
        var result = new List<OfficeGradientStop>(stops.Count * 2 - 1);
        // Positive focus places the last color inside the frame; negative focus
        // places the first color inside it. Both sides retain the source ramp.
        double pivot = focus > 0 ? (100 - focus) / 100D : -focus / 100D;
        if (focus > 0) {
            result.AddRange(stops.Select(stop => new OfficeGradientStop(stop.Offset * pivot, stop.Color)));
            result.AddRange(stops.AsEnumerable().Reverse().Skip(1)
                .Select(stop => new OfficeGradientStop(pivot + (1 - stop.Offset) * (1 - pivot), stop.Color)));
        } else {
            result.AddRange(stops.AsEnumerable().Reverse().Select(stop => new OfficeGradientStop((1 - stop.Offset) * pivot, stop.Color)));
            result.AddRange(stops.Skip(1).Select(stop => new OfficeGradientStop(pivot + stop.Offset * (1 - pivot), stop.Color)));
        }
        return result;
    }
}

internal sealed class OfficeArtLinearGradientFill {
    internal OfficeArtLinearGradientFill(OfficeLinearGradient gradient, double opacity, bool opacityApproximated) {
        Gradient = gradient; Opacity = opacity; OpacityApproximated = opacityApproximated;
    }
    internal OfficeLinearGradient Gradient { get; }
    internal double Opacity { get; }
    internal bool OpacityApproximated { get; }
}
