using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private readonly struct ChartPlotBounds {
        internal ChartPlotBounds(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }

        internal bool Contains(OfficePoint point) =>
            point.X >= Left && point.X <= Left + Width &&
            point.Y >= Top && point.Y <= Top + Height;
    }

    private static bool HasExplicitValueBounds(OfficeChartLayout layout, OfficeChartAxisGroup? axisGroup,
        bool horizontalValueAxis = false) =>
        axisGroup == OfficeChartAxisGroup.Secondary
            ? layout.SecondaryValueAxis?.Minimum.HasValue == true || layout.SecondaryValueAxis?.Maximum.HasValue == true
            : horizontalValueAxis
                ? layout.HorizontalAxisMinimum.HasValue || layout.HorizontalAxisMaximum.HasValue
                : layout.VerticalAxisMinimum.HasValue || layout.VerticalAxisMaximum.HasValue;

    private static bool TryClipValueSegment(OfficePoint from, OfficePoint to, ValueRange range,
        ChartPlotBounds bounds, out OfficePoint start, out OfficePoint end) {
        start = default;
        end = default;
        double scale = Math.Max(1D, Math.Max(Math.Max(Math.Abs(from.Y), Math.Abs(to.Y)),
            Math.Max(Math.Abs(range.Min), Math.Abs(range.Max))));
        double y0 = from.Y / scale, y1 = to.Y / scale;
        double lower = range.Min / scale, upper = range.Max / scale;
        double delta = y1 - y0;
        double entry = 0D, exit = 1D;
        if (!Clip(-delta, y0 - lower, ref entry, ref exit) ||
            !Clip(delta, upper - y0, ref entry, ref exit)) return false;
        double width = to.X - from.X;
        double firstValue = entry == 0D ? from.Y : from.Y < range.Min ? range.Min : range.Max;
        double lastValue = exit == 1D ? to.Y : to.Y < range.Min ? range.Min : range.Max;
        start = new OfficePoint(from.X + width * entry,
            ToPlotY(firstValue, range.Min, range.Max, bounds.Top, bounds.Height));
        end = new OfficePoint(from.X + width * exit,
            ToPlotY(lastValue, range.Min, range.Max, bounds.Top, bounds.Height));
        return true;
    }

    private static List<OfficePoint> ClipValuePolygon(List<OfficePoint> polygon,
        ValueRange range, ChartPlotBounds bounds) {
        polygon = ClipValueEdge(polygon, range.Min, keepGreater: true);
        polygon = ClipValueEdge(polygon, range.Max, keepGreater: false);
        for (int index = 0; index < polygon.Count; index++) {
            OfficePoint point = polygon[index];
            polygon[index] = new OfficePoint(point.X,
                ToPlotY(point.Y, range.Min, range.Max, bounds.Top, bounds.Height));
        }
        return polygon;
    }

    private static List<OfficePoint> ClipValueEdge(List<OfficePoint> polygon,
        double edge, bool keepGreater) {
        var clipped = new List<OfficePoint>();
        if (polygon.Count == 0) return clipped;
        OfficePoint previous = polygon[polygon.Count - 1];
        bool previousInside = keepGreater ? previous.Y >= edge : previous.Y <= edge;
        foreach (OfficePoint current in polygon) {
            bool currentInside = keepGreater ? current.Y >= edge : current.Y <= edge;
            if (currentInside != previousInside) {
                double scale = Math.Max(1D, Math.Max(Math.Abs(previous.Y),
                    Math.Max(Math.Abs(current.Y), Math.Abs(edge))));
                double from = previous.Y / scale, to = current.Y / scale;
                double fraction = (edge / scale - from) / (to - from);
                clipped.Add(new OfficePoint(previous.X + (current.X - previous.X) * fraction, edge));
            }
            if (currentInside) clipped.Add(current);
            previous = current;
            previousInside = currentInside;
        }
        return clipped;
    }

    private static void AddClippedValueLine(OfficeDrawing drawing,
        IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth,
        OfficeStrokeDashStyle dashStyle, ValueRange range, ChartPlotBounds bounds) {
        for (int index = 1; index < points.Count; index++)
            if (TryClipValueSegment(points[index - 1], points[index], range, bounds,
                    out OfficePoint start, out OfficePoint end) &&
                (start.X != end.X || start.Y != end.Y))
                AddPointLine(drawing, new[] { start, end }, color, strokeWidth, dashStyle);
    }

    private static void AddScatterPlotLine(OfficeDrawing drawing, System.Collections.Generic.IReadOnlyList<OfficePoint> points,
        OfficeColor color, double strokeWidth, OfficeStrokeDashStyle dashStyle, bool clipPlot,
        ValueRange xRange, ValueRange yRange, ChartPlotBounds bounds) {
        if (!clipPlot) {
            AddPointLine(drawing, points, color, strokeWidth, dashStyle);
            return;
        }
        for (int index = 1; index < points.Count; index++) {
            if (!TryClipScatterSegment(points[index - 1], points[index], xRange, yRange, bounds,
                    out OfficePoint start, out OfficePoint end)) continue;
            AddPointLine(drawing, new[] { start, end }, color, strokeWidth, dashStyle);
        }
    }

    private static bool TryClipScatterSegment(OfficePoint from, OfficePoint to, ValueRange xRange,
        ValueRange yRange, ChartPlotBounds bounds, out OfficePoint start, out OfficePoint end) {
        start = default;
        end = default;
        // Parameter rounding loses the visible endpoint when a very distant point
        // is first: the entry fraction can round all the way to one. Clip from
        // the in-range endpoint so the small fraction remains representable.
        bool fromInside = from.X >= xRange.Min && from.X <= xRange.Max &&
            from.Y >= yRange.Min && from.Y <= yRange.Max;
        bool toInside = to.X >= xRange.Min && to.X <= xRange.Max &&
            to.Y >= yRange.Min && to.Y <= yRange.Max;
        if (!fromInside && toInside) {
            if (!TryClipScatterSegment(to, from, xRange, yRange, bounds,
                    out OfficePoint reversedStart, out OfficePoint reversedEnd)) return false;
            start = reversedEnd;
            end = reversedStart;
            return true;
        }
        double x0 = (from.X - xRange.Min) / (xRange.Max - xRange.Min);
        double y0 = (from.Y - yRange.Min) / (yRange.Max - yRange.Min);
        double x1 = (to.X - xRange.Min) / (xRange.Max - xRange.Min);
        double y1 = (to.Y - yRange.Min) / (yRange.Max - yRange.Min);
        if (double.IsNaN(x0) || double.IsNaN(y0) || double.IsNaN(x1) || double.IsNaN(y1) ||
            double.IsInfinity(x0) || double.IsInfinity(y0) || double.IsInfinity(x1) || double.IsInfinity(y1))
            return false;
        double dx = x1 - x0;
        double dy = y1 - y0;
        if (double.IsInfinity(dx) || double.IsInfinity(dy)) return false;
        double entry = 0D;
        double exit = 1D;
        if (!Clip(-dx, x0, ref entry, ref exit) || !Clip(dx, 1D - x0, ref entry, ref exit) ||
            !Clip(-dy, y0, ref entry, ref exit) || !Clip(dy, 1D - y0, ref entry, ref exit))
            return false;
        start = Map(x0 + entry * dx, y0 + entry * dy);
        end = Map(x0 + exit * dx, y0 + exit * dy);
        return true;

        OfficePoint Map(double x, double y) => new OfficePoint(
            bounds.Left + bounds.Width * Math.Max(0D, Math.Min(1D, x)),
            bounds.Top + bounds.Height * (1D - Math.Max(0D, Math.Min(1D, y))));
    }

    private static bool Clip(double p, double q, ref double entry, ref double exit) {
        if (p == 0D) return q >= 0D;
        double edge = q / p;
        if (p < 0D) {
            if (edge > exit) return false;
            if (edge > entry) entry = edge;
        } else {
            if (edge < entry) return false;
            if (edge < exit) exit = edge;
        }
        return true;
    }

    private static void AddClippedPlotGeometry(OfficeDrawing drawing, OfficeDrawing geometry,
        ChartPlotBounds bounds) {
        drawing.AddClippedDrawing(geometry, bounds.Left, bounds.Top,
            OfficeClipPath.Rectangle(bounds.Width, bounds.Height), -bounds.Left, -bounds.Top);
    }

    private static void AddClippedPlotGeometry(OfficeDrawing drawing, OfficeDrawing geometry,
        OfficeDrawing labels, ChartPlotBounds bounds) {
        AddClippedPlotGeometry(drawing, geometry, bounds);
        drawing.AddDrawing(labels, 0D, 0D);
    }
}
