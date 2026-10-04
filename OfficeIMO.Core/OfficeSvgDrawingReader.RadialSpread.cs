using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed partial class SvgGradientDefinition {
        private bool TryCreateRadialSpread(OfficeRadialGradient field, OfficeShape shape, out OfficeRadialGradient? gradient) {
            gradient = field;
            if (SpreadMode == SvgGradientSpreadMode.Pad) return true;
            // A point focus strictly inside the end ellipse gives nested level sets.
            // Their gauge is convex, so the paint rectangle's corners bound every
            // needed cycle. Nonzero focal circles need a separate inner-domain policy.
            if (field.StartRadiusX != 0D || field.StartRadiusY != 0D) return false;
            double dx = (field.StartX - field.EndX) / field.EndRadiusX;
            double dy = (field.StartY - field.EndY) / field.EndRadiusY;
            if (dx * dx + dy * dy >= 1D) return false;
            double margin = shape.StrokeWidth * Math.Max(1D, shape.StrokeMiterLimit);
            double padX = margin / shape.Width, padY = margin / shape.Height;
            double maximum = 1D;
            foreach (var point in new[] { new OfficePoint(-padX, -padY), new OfficePoint(1D + padX, -padY),
                new OfficePoint(1D + padX, 1D + padY), new OfficePoint(-padX, 1D + padY) }) {
                maximum = Math.Max(maximum, field.SampleUnboundedRatio(point.X, point.Y));
            }
            maximum = Math.Ceiling(maximum);
            if (double.IsNaN(maximum) || double.IsInfinity(maximum) || maximum > MaximumGradientStops) return false;
            var expanded = new List<SvgExpandedGradientStop> { new SvgExpandedGradientStop(0D, Stops[0].Color) };
            int cycles = (int)maximum;
            for (int cycle = 0; cycle < cycles; cycle++) {
                AddExpandedStops(cycle, SpreadMode == SvgGradientSpreadMode.Reflect && (cycle & 1) != 0, 0D, maximum, expanded);
                if (expanded.Count > MaximumGradientStops) return false;
            }
            expanded.Add(new SvgExpandedGradientStop(maximum, SpreadMode == SvgGradientSpreadMode.Reflect && (cycles & 1) == 0 ? Stops[0].Color : Stops[Stops.Count - 1].Color));
            if (expanded.Count > MaximumGradientStops) return false;
            var stops = expanded.OrderBy(stop => stop.Position)
                .Select(stop => new OfficeGradientStop(stop.Position / maximum, stop.Color)).ToArray();
            gradient = new OfficeRadialGradient(field.StartX, field.StartY, 0D, 0D,
                field.StartX + maximum * (field.EndX - field.StartX), field.StartY + maximum * (field.EndY - field.StartY),
                field.EndRadiusX * maximum, field.EndRadiusY * maximum, stops).TransformCoordinates(field.CoordinateTransform);
            return true;
        }
    }
}
