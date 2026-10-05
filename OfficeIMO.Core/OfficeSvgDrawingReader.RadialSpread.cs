using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed partial class SvgGradientDefinition {
        private bool TryCreateRadialSpread(OfficeRadialGradient field, OfficeShape shape, out OfficeRadialGradient? gradient) {
            gradient = field;
            if (SpreadMode == SvgGradientSpreadMode.Pad) return true;
            double margin = shape.StrokeWidth * Math.Max(1D, shape.StrokeMiterLimit);
            double padX = margin / shape.Width, padY = margin / shape.Height;
            var corners = new[] { new OfficePoint(-padX, -padY), new OfficePoint(1D + padX, -padY),
                new OfficePoint(1D + padX, 1D + padY), new OfficePoint(-padX, 1D + padY) };
            if (!field.TryGetSpreadCycles(corners, UseFirstRadialIntersection,
                SpreadMode == SvgGradientSpreadMode.Reflect, MaximumGradientStops, out int cycles)) return false;
            double maximum = cycles;
            double dx = (field.StartX - field.EndX) / field.EndRadiusX;
            double dy = (field.StartY - field.EndY) / field.EndRadiusY;
            bool nativeExterior = dx * dx + dy * dy >= 1D;
            var expanded = new List<SvgExpandedGradientStop> { new SvgExpandedGradientStop(0D, Stops[0].Color) };
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
            if (nativeExterior) gradient = gradient.WithFirstPadIntersection();
            return true;
        }
    }
}
