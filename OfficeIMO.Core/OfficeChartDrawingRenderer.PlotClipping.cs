using System;

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

    private static bool HasExplicitValueBounds(OfficeChartLayout layout, OfficeChartAxisGroup? axisGroup) =>
        axisGroup == OfficeChartAxisGroup.Secondary
            ? layout.SecondaryValueAxis?.Minimum.HasValue == true || layout.SecondaryValueAxis?.Maximum.HasValue == true
            : layout.VerticalAxisMinimum.HasValue || layout.VerticalAxisMaximum.HasValue;

    private static double ToUnclampedPlotY(double value, double min, double max, double plotTop, double plotHeight) =>
        plotTop + plotHeight * (1D - BoundedPlotRatio(value, min, max));

    private static double ToUnclampedPlotX(double value, double min, double max, double plotLeft, double plotWidth) =>
        plotLeft + plotWidth * BoundedPlotRatio(value, min, max);

    private static double BoundedPlotRatio(double value, double min, double max) {
        double range = max - min;
        if (range <= 0D || double.IsInfinity(range)) return 0.5D;
        double ratio = (value - min) / range;
        if (double.IsNaN(ratio)) return 0.5D;
        return Math.Max(-1000000D, Math.Min(1000000D, ratio));
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
