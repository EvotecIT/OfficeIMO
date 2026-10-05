using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed partial class SvgGradientDefinition {
        private bool TryCreateRadialSpread(OfficeRadialGradient field, OfficeShape shape, out OfficeRadialGradient? gradient) {
            gradient = field;
            if (SpreadMode == SvgGradientSpreadMode.Pad) return true;
            // Interior point foci have a convex gauge. Native exterior foci instead
            // use the first physical root, bounded by the square root of the root
            // product C/A. C is convex, so its maximum also occurs at a box corner.
            // Boundary foci have A=0 and may require infinitely many cycles.
            if (field.StartRadiusX != 0D || field.StartRadiusY != 0D) return false;
            double dx = (field.StartX - field.EndX) / field.EndRadiusX;
            double dy = (field.StartY - field.EndY) / field.EndRadiusY;
            double a = dx * dx + dy * dy - 1D;
            bool exterior = a > 0D;
            if (a >= 0D && (!exterior || !UseFirstRadialIntersection)) return false;
            double margin = shape.StrokeWidth * Math.Max(1D, shape.StrokeMiterLimit);
            double padX = margin / shape.Width, padY = margin / shape.Height;
            double maximum = 1D;
            foreach (var point in new[] { new OfficePoint(-padX, -padY), new OfficePoint(1D + padX, -padY),
                new OfficePoint(1D + padX, 1D + padY), new OfficePoint(-padX, 1D + padY) }) {
                double bound;
                if (exterior) {
                    var local = field.CoordinateTransform.Invert().TransformPoint(point);
                    double px = (local.X - field.StartX) / field.EndRadiusX;
                    double py = (local.Y - field.StartY) / field.EndRadiusY;
                    bound = Math.Sqrt((px * px + py * py) / a);
                } else bound = field.SampleUnboundedRatio(point.X, point.Y);
                maximum = Math.Max(maximum, bound);
            }
            maximum = Math.Ceiling(maximum);
            // Native Reflect paints offset zero outside the cone. Ending on an
            // even cycle makes that color the reversed field's exterior endpoint.
            if (exterior && SpreadMode == SvgGradientSpreadMode.Reflect && maximum % 2D != 0D) maximum++;
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
            if (exterior) gradient = gradient.WithFirstPadIntersection();
            return true;
        }
    }
}
