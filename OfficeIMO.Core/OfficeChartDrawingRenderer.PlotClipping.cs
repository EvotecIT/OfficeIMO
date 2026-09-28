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

    private static double ToUnclampedPlotY(double value, double min, double max, double plotTop, double plotHeight) =>
        FinitePlotCoordinate(plotTop + plotHeight * (1D - UnclampedPlotRatio(value, min, max)));

    private static double ToUnclampedPlotX(double value, double min, double max, double plotLeft, double plotWidth) =>
        FinitePlotCoordinate(plotLeft + plotWidth * UnclampedPlotRatio(value, min, max));

    private static double UnclampedPlotRatio(double value, double min, double max) {
        double range = max - min;
        if (range <= 0D || double.IsInfinity(range)) return 0.5D;
        double ratio = (value - min) / range;
        if (double.IsNaN(ratio)) return 0.5D;
        if (double.IsInfinity(ratio))
            throw new NotSupportedException("The chart value is too far outside the explicit plot bounds to render faithfully.");
        return ratio;
    }

    private static double FinitePlotCoordinate(double coordinate) {
        if (double.IsNaN(coordinate) || double.IsInfinity(coordinate))
            throw new NotSupportedException("The chart value is too far outside the explicit plot bounds to render faithfully.");
        return coordinate;
    }

    private static bool TryClipPlotSegment(OfficePoint from, OfficePoint to, ChartPlotBounds bounds,
        out OfficePoint start, out OfficePoint end) {
        start = default;
        end = default;
        double x0 = (from.X - bounds.Left) / bounds.Width;
        double y0 = (from.Y - bounds.Top) / bounds.Height;
        double x1 = (to.X - bounds.Left) / bounds.Width;
        double y1 = (to.Y - bounds.Top) / bounds.Height;
        double dx = x1 - x0, dy = y1 - y0;
        if (double.IsNaN(dx) || double.IsNaN(dy) || double.IsInfinity(dx) || double.IsInfinity(dy))
            throw new NotSupportedException("The chart segment cannot be clipped faithfully at these values.");
        double entry = 0D, exit = 1D;
        if (!Clip(-dx, x0, ref entry, ref exit) || !Clip(dx, 1D - x0, ref entry, ref exit) ||
            !Clip(-dy, y0, ref entry, ref exit) || !Clip(dy, 1D - y0, ref entry, ref exit)) return false;
        start = new OfficePoint(bounds.Left + bounds.Width * Math.Max(0D, Math.Min(1D, x0 + entry * dx)),
            bounds.Top + bounds.Height * Math.Max(0D, Math.Min(1D, y0 + entry * dy)));
        end = new OfficePoint(bounds.Left + bounds.Width * Math.Max(0D, Math.Min(1D, x0 + exit * dx)),
            bounds.Top + bounds.Height * Math.Max(0D, Math.Min(1D, y0 + exit * dy)));
        return true;
    }

    private static void AddClippedPointLine(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points,
        OfficeColor color, double strokeWidth, OfficeStrokeDashStyle dashStyle, ChartPlotBounds bounds) {
        for (int index = 1; index < points.Count; index++)
            if (TryClipPlotSegment(points[index - 1], points[index], bounds,
                    out OfficePoint start, out OfficePoint end) &&
                (start.X != end.X || start.Y != end.Y))
                AddPointLine(drawing, new[] { start, end }, color, strokeWidth, dashStyle);
    }

    private static List<OfficePoint> ClipPlotPolygon(List<OfficePoint> polygon, ChartPlotBounds bounds) {
        polygon = ClipEdge(polygon, bounds.Left, vertical: true, keepGreater: true);
        polygon = ClipEdge(polygon, bounds.Left + bounds.Width, vertical: true, keepGreater: false);
        polygon = ClipEdge(polygon, bounds.Top, vertical: false, keepGreater: true);
        return ClipEdge(polygon, bounds.Top + bounds.Height, vertical: false, keepGreater: false);
    }

    private static List<OfficePoint> ClipEdge(List<OfficePoint> polygon, double edge,
        bool vertical, bool keepGreater) {
        var clipped = new List<OfficePoint>();
        if (polygon.Count == 0) return clipped;
        OfficePoint previous = polygon[polygon.Count - 1];
        bool previousInside = Inside(previous);
        foreach (OfficePoint current in polygon) {
            bool currentInside = Inside(current);
            if (currentInside != previousInside) clipped.Add(Intersection(previous, current));
            if (currentInside) clipped.Add(current);
            previous = current;
            previousInside = currentInside;
        }
        return clipped;

        bool Inside(OfficePoint point) => keepGreater
            ? (vertical ? point.X : point.Y) >= edge
            : (vertical ? point.X : point.Y) <= edge;
        OfficePoint Intersection(OfficePoint from, OfficePoint to) {
            double fromCoordinate = vertical ? from.X : from.Y;
            double toCoordinate = vertical ? to.X : to.Y;
            double difference = toCoordinate - fromCoordinate;
            if (double.IsInfinity(difference))
                throw new NotSupportedException("The chart area cannot be clipped faithfully at these values.");
            double fraction = (edge - fromCoordinate) / difference;
            return vertical
                ? new OfficePoint(edge, from.Y + fraction * (to.Y - from.Y))
                : new OfficePoint(from.X + fraction * (to.X - from.X), edge);
        }
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
