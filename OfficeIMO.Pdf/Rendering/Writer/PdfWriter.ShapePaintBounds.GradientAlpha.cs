using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Clips diagnostic geometry to nonzero alpha intervals of its actual gradient field.</summary>
    private static bool TryCreateGradientAlphaContours(OfficeShape shape, List<List<OfficePoint>> paint,
        OfficeLinearGradient? linear, OfficeRadialGradient? radial, out List<List<OfficePoint>>? contours, Action<int>? chargePoints = null) {
        contours = null;
        IReadOnlyList<OfficeGradientStop>? stops = radial?.Stops ?? linear?.Stops;
        if (stops == null) return true;
        chargePoints?.Invoke(stops.Count);
        stops = PdfVisualResourceDictionaryBuilder.NormalizeGradientStops(stops);
        bool transparent = false, visible = false;
        foreach (var stop in stops) { transparent |= stop.Color.A == 0; visible |= stop.Color.A > 0; }
        if (!transparent) return true;
        if (!visible && radial?.OutsideColor?.A is not > 0) { contours = new(); return true; }

        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (var ring in paint) foreach (OfficePoint point in ring) {
            left = Math.Min(left, point.X); right = Math.Max(right, point.X);
            top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
        }
        contours = new();
        if (!(right > left && bottom > top)) return true;
        var rectangle = new List<OfficePoint> { new(left, top), new(right, top), new(right, bottom), new(left, bottom) };
        // Non-nested/moving radial cones and repeated alpha fields need a
        // separate native-field intersection; do not label their whole canvas as ink.
        if (radial != null && (radial.SpreadMode != OfficeGradientSpreadMode.Pad || radial.OutsideColor != null
            || radial.StartX != radial.EndX || radial.StartY != radial.EndY
            || radial.EndRadiusX <= radial.StartRadiusX || radial.EndRadiusY <= radial.StartRadiusY)) return false;

        for (int index = 0; index < stops.Count - 1; index++) {
            if (stops[index].Color.A == 0 && stops[index + 1].Color.A == 0) continue;
            double lower = index == 0 && stops[0].Color.A > 0 ? double.NegativeInfinity : stops[index].Offset;
            int last = index;
            while (last + 2 < stops.Count && (stops[last + 1].Color.A > 0 || stops[last + 2].Color.A > 0)) last++;
            double upper = last == stops.Count - 2 && stops[stops.Count - 1].Color.A > 0
                ? double.PositiveInfinity : stops[last + 1].Offset;
            index = last;
            if (radial != null) {
                if (double.IsPositiveInfinity(upper)) { chargePoints?.Invoke(rectangle.Count); contours.Add(new(rectangle)); }
                else contours.Add(RadialRing(upper));
                if (!double.IsNegativeInfinity(lower) &&
                    (radial.StartRadiusX + (radial.EndRadiusX - radial.StartRadiusX) * lower) > 0)
                    contours.Add(RadialRing(lower));
            } else {
                var polygon = new List<OfficePoint>(rectangle);
                if (!double.IsNegativeInfinity(lower)) polygon = ClipRatio(polygon, lower, true);
                if (!double.IsPositiveInfinity(upper)) polygon = ClipRatio(polygon, upper, false);
                if (polygon.Count >= 3) { chargePoints?.Invoke(polygon.Count); contours.Add(polygon); }
            }
        }
        return true;

        double Ratio(OfficePoint p) {
            double dx = linear!.EndX - linear.StartX, dy = linear.EndY - linear.StartY;
            return ((p.X / Math.Max(.0001, shape.Width) - linear.StartX) * dx
                + (p.Y / Math.Max(.0001, shape.Height) - linear.StartY) * dy) / (dx * dx + dy * dy);
        }
        List<OfficePoint> ClipRatio(List<OfficePoint> polygon, double boundary, bool above) {
            var result = new List<OfficePoint>();
            if (polygon.Count == 0) return result;
            OfficePoint previous = polygon[polygon.Count - 1];
            double previousDistance = (Ratio(previous) - boundary) * (above ? 1 : -1);
            foreach (OfficePoint current in polygon) {
                double distance = (Ratio(current) - boundary) * (above ? 1 : -1);
                if ((previousDistance >= 0) != (distance >= 0)) {
                    double t = previousDistance / (previousDistance - distance);
                    result.Add(new(previous.X + t * (current.X - previous.X), previous.Y + t * (current.Y - previous.Y)));
                }
                if (distance >= 0) result.Add(current);
                previous = current; previousDistance = distance;
            }
            return result;
        }
        List<OfficePoint> RadialRing(double ratio) {
            double rx = radial!.StartRadiusX + (radial.EndRadiusX - radial.StartRadiusX) * ratio;
            double ry = radial.StartRadiusY + (radial.EndRadiusY - radial.StartRadiusY) * ratio;
            OfficeTransform fieldToPage = radial.CoordinateTransform
                .Then(OfficeTransform.Scale(Math.Max(.0001, shape.Width), Math.Max(.0001, shape.Height)))
                .Then(shape.Transform ?? OfficeTransform.Identity);
            double density = OfficeStrokeGeometry.LocalFlatteningDensity(fieldToPage, 4D);
            var ring = OfficeCurveFlattening.Ellipse(radial.StartX, radial.StartY, rx, ry, density);
            chargePoints?.Invoke(ring.Count);
            for (int i = 0; i < ring.Count; i++) {
                OfficePoint normalized = radial.CoordinateTransform.TransformPoint(ring[i]);
                ring[i] = new(normalized.X * Math.Max(.0001, shape.Width), normalized.Y * Math.Max(.0001, shape.Height));
            }
            return ring;
        }
    }
}
