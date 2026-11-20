namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const double ErrorBarCapHalfLength = 3.5D;

    /// <summary>
    /// Draws a point's error bars along its value axis. <paramref name="offset"/> moves the bar to the point's
    /// stacked position; <paramref name="xValues"/> selects the x-direction bars of a scatter point.
    /// </summary>
    private static void AddPointErrorBars(OfficeDrawing drawing, OfficeSeries series, int index, double value, double offset,
        double crossPixel, bool horizontal, Func<double, double> toPixel, ValueRange range, bool xValues = false,
        IReadOnlyList<double>? spanValues = null) {
        if (series.ErrorBars == null) return;
        foreach (OfficeChartErrorBars bars in series.ErrorBars) {
            if (bars.AlongCategoryAxis != xValues) continue;
            IReadOnlyList<double> values = spanValues ?? (xValues ? series.XValues ?? Array.Empty<double>() : series.Values);
            if (!bars.TryGetSpan(values, index, value, out double center, out double low, out double high)) continue;
            double clampedCenter = ClampValueToRange(center + offset, range.Min, range.Max);
            double clampedLow = ClampValueToRange(low + offset, range.Min, range.Max);
            double clampedHigh = ClampValueToRange(high + offset, range.Min, range.Max);
            double from = toPixel(clampedLow), to = toPixel(clampedHigh), middle = toPixel(clampedCenter);
            if (!IsFiniteChartValue(from) || !IsFiniteChartValue(to) || !IsFiniteChartValue(crossPixel)) continue;
            OfficeColor color = bars.Color ?? OfficeColor.FromRgb(0x59, 0x59, 0x59);
            double width = bars.Width ?? 0.75D;
            OfficeStrokeDashStyle dash = bars.DashStyle ?? OfficeStrokeDashStyle.Solid;
            if (Math.Abs(to - from) > 0.01D) AddErrorBarLine(drawing, horizontal, crossPixel, from, to, color, width, dash);
            if (!bars.EndCaps) continue;
            // A cap closes each end the bar reaches; an end cut off by the axis bounds has no cap.
            if (bars.Direction != OfficeChartErrorBarDirection.Minus && clampedHigh == high + offset && Math.Abs(to - middle) > 0.01D)
                AddErrorBarLine(drawing, !horizontal, to, crossPixel - ErrorBarCapHalfLength, crossPixel + ErrorBarCapHalfLength, color, width, OfficeStrokeDashStyle.Solid);
            if (bars.Direction != OfficeChartErrorBarDirection.Plus && clampedLow == low + offset && Math.Abs(from - middle) > 0.01D)
                AddErrorBarLine(drawing, !horizontal, from, crossPixel - ErrorBarCapHalfLength, crossPixel + ErrorBarCapHalfLength, color, width, OfficeStrokeDashStyle.Solid);
        }
    }

    // A horizontal line runs along x at the given y; a vertical line runs along y at the given x.
    private static void AddErrorBarLine(OfficeDrawing drawing, bool horizontal, double cross, double start, double end,
        OfficeColor color, double width, OfficeStrokeDashStyle dash) {
        OfficePoint first = horizontal ? new OfficePoint(start, cross) : new OfficePoint(cross, start);
        OfficePoint second = horizontal ? new OfficePoint(end, cross) : new OfficePoint(cross, end);
        AddPointLine(drawing, new[] { first, second }, color, width, dash);
    }
}
