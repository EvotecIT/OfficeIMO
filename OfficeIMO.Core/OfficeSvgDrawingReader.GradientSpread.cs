using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed partial class SvgGradientDefinition {
        private bool TryCreateLinearSpread(
            double x1,
            double y1,
            double x2,
            double y2,
            out OfficeLinearGradient? gradient) {
            gradient = null;
            if (SpreadMode == SvgGradientSpreadMode.Pad) {
                gradient = OfficeLinearGradient.CreateImported(x1, y1, x2, y2, Stops);
                return true;
            }

            double dx = x2 - x1;
            double dy = y2 - y1;
            double lengthSquared = (dx * dx) + (dy * dy);
            if (lengthSquared <= 0.000000000001D) return false;
            double minimum = double.PositiveInfinity;
            double maximum = double.NegativeInfinity;
            IncludeLinearRatio(0D, 0D, x1, y1, dx, dy, lengthSquared, ref minimum, ref maximum);
            IncludeLinearRatio(1D, 0D, x1, y1, dx, dy, lengthSquared, ref minimum, ref maximum);
            IncludeLinearRatio(0D, 1D, x1, y1, dx, dy, lengthSquared, ref minimum, ref maximum);
            IncludeLinearRatio(1D, 1D, x1, y1, dx, dy, lengthSquared, ref minimum, ref maximum);
            if (maximum - minimum <= 0.000000000001D) return false;

            var expanded = new List<SvgExpandedGradientStop> {
                new SvgExpandedGradientStop(minimum, EvaluateSpreadColor(minimum))
            };
            double cycleDivisor = SpreadMode == SvgGradientSpreadMode.Reflect ? 2D : 1D;
            double firstCycleValue = Math.Floor(minimum / cycleDivisor) - 1D;
            double lastCycleValue = Math.Ceiling(maximum / cycleDivisor) + 1D;
            if (firstCycleValue < int.MinValue || lastCycleValue > int.MaxValue) return false;
            int firstCycle = (int)firstCycleValue;
            int lastCycle = (int)lastCycleValue;
            if (lastCycle - firstCycle > MaximumGradientStops) return false;
            for (int cycle = firstCycle; cycle <= lastCycle; cycle++) {
                if (SpreadMode == SvgGradientSpreadMode.Repeat) {
                    AddExpandedStops(cycle, reverse: false, minimum, maximum, expanded);
                } else {
                    AddExpandedStops(cycle * 2D, reverse: false, minimum, maximum, expanded);
                    AddExpandedStops((cycle * 2D) + 1D, reverse: true, minimum, maximum, expanded);
                }
                if (expanded.Count > MaximumGradientStops) return false;
            }
            expanded.Add(new SvgExpandedGradientStop(maximum, EvaluateSpreadColor(maximum)));
            expanded = expanded.OrderBy(stop => stop.Position).ToList();
            if (expanded.Count > MaximumGradientStops) return false;

            double span = maximum - minimum;
            var concreteStops = new List<OfficeGradientStop>(expanded.Count);
            for (int index = 0; index < expanded.Count; index++) {
                SvgExpandedGradientStop stop = expanded[index];
                concreteStops.Add(new OfficeGradientStop((stop.Position - minimum) / span, stop.Color));
            }
            gradient = OfficeLinearGradient.CreateImported(
                x1 + (dx * minimum),
                y1 + (dy * minimum),
                x1 + (dx * maximum),
                y1 + (dy * maximum),
                concreteStops);
            return true;
        }

        private void AddExpandedStops(
            double cycleStart,
            bool reverse,
            double minimum,
            double maximum,
            ICollection<SvgExpandedGradientStop> expanded) {
            if (!reverse) {
                for (int index = 0; index < Stops.Count; index++) {
                    AddExpandedStop(cycleStart + Stops[index].Offset, Stops[index].Color, minimum, maximum, expanded);
                }
                return;
            }
            for (int index = Stops.Count - 1; index >= 0; index--) {
                AddExpandedStop(cycleStart + (1D - Stops[index].Offset), Stops[index].Color, minimum, maximum, expanded);
            }
        }

        private static void AddExpandedStop(
            double position,
            OfficeColor color,
            double minimum,
            double maximum,
            ICollection<SvgExpandedGradientStop> expanded) {
            if (position >= minimum && position <= maximum) expanded.Add(new SvgExpandedGradientStop(position, color));
        }

        private OfficeColor EvaluateSpreadColor(double ratio) {
            double mapped;
            if (SpreadMode == SvgGradientSpreadMode.Repeat) {
                mapped = ratio - Math.Floor(ratio);
            } else {
                mapped = ratio % 2D;
                if (mapped < 0D) mapped += 2D;
                if (mapped > 1D) mapped = 2D - mapped;
            }
            if (mapped <= Stops[0].Offset) return Stops[0].Color;
            for (int index = 1; index < Stops.Count; index++) {
                OfficeGradientStop next = Stops[index];
                if (mapped <= next.Offset) {
                    OfficeGradientStop previous = Stops[index - 1];
                    double span = next.Offset - previous.Offset;
                    double amount = span <= double.Epsilon ? 0D : (mapped - previous.Offset) / span;
                    return InterpolateColor(previous.Color, next.Color, amount);
                }
            }
            return Stops[Stops.Count - 1].Color;
        }

        private static OfficeColor InterpolateColor(OfficeColor start, OfficeColor end, double amount) {
            double clamped = amount < 0D ? 0D : amount > 1D ? 1D : amount;
            return OfficeColor.FromRgba(
                InterpolateByte(start.R, end.R, clamped),
                InterpolateByte(start.G, end.G, clamped),
                InterpolateByte(start.B, end.B, clamped),
                InterpolateByte(start.A, end.A, clamped));
        }

        private static byte InterpolateByte(byte start, byte end, double amount) =>
            (byte)Math.Round(start + ((end - start) * amount));

        private static void IncludeLinearRatio(
            double x,
            double y,
            double startX,
            double startY,
            double dx,
            double dy,
            double lengthSquared,
            ref double minimum,
            ref double maximum) {
            double ratio = (((x - startX) * dx) + ((y - startY) * dy)) / lengthSquared;
            minimum = Math.Min(minimum, ratio);
            maximum = Math.Max(maximum, ratio);
        }
    }

    private readonly struct SvgExpandedGradientStop {
        internal double Position { get; }
        internal OfficeColor Color { get; }

        internal SvgExpandedGradientStop(double position, OfficeColor color) {
            Position = position;
            Color = color;
        }
    }
}
