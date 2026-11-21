namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const double ErrorBarCapHalfLength = 3.5D;

    /// <summary>
    /// Draws a point's error bars along its value axis. <paramref name="value"/> and <paramref name="offset"/> are
    /// plotted values: the point and how far stacking moves it. <paramref name="xValues"/> selects the x-direction bars
    /// of a scatter point.
    /// </summary>
    private static void AddPointErrorBars(OfficeDrawing drawing, OfficeSeries series, int index, double value, double offset,
        double crossPixel, bool horizontal, Func<double, double> toPixel, ValueRange range, bool xValues = false,
        IReadOnlyList<double>? spanValues = null) {
        if (series.ErrorBars == null) return;
        foreach (OfficeChartErrorBars bars in series.ErrorBars) {
            if (bars.AlongCategoryAxis != xValues) continue;
            if (!TryGetPlottedErrorSpan(series, bars, index, value, offset, xValues, spanValues,
                    out double center, out double minusEnd, out double plusEnd)) continue;
            double clampedCenter = ClampValueToRange(center, range.Min, range.Max);
            double clampedMinus = ClampValueToRange(minusEnd, range.Min, range.Max);
            double clampedPlus = ClampValueToRange(plusEnd, range.Min, range.Max);
            double from = toPixel(clampedMinus), to = toPixel(clampedPlus), middle = toPixel(clampedCenter);
            if (!IsFiniteChartValue(from) || !IsFiniteChartValue(to) || !IsFiniteChartValue(crossPixel)) continue;
            OfficeColor color = bars.Color ?? OfficeColor.FromRgb(0x59, 0x59, 0x59);
            double width = bars.Width ?? 0.75D;
            OfficeStrokeDashStyle dash = bars.DashStyle ?? OfficeStrokeDashStyle.Solid;
            if (Math.Abs(to - from) > 0.01D) AddErrorBarLine(drawing, horizontal, crossPixel, from, to, color, width, dash);
            if (!bars.EndCaps) continue;
            // A cap closes each end the bar reaches; an end cut off by the axis bounds has no cap.
            if (bars.Direction != OfficeChartErrorBarDirection.Minus && clampedPlus == plusEnd && Math.Abs(to - middle) > 0.01D)
                AddErrorBarLine(drawing, !horizontal, to, crossPixel - ErrorBarCapHalfLength, crossPixel + ErrorBarCapHalfLength, color, width, OfficeStrokeDashStyle.Solid);
            if (bars.Direction != OfficeChartErrorBarDirection.Plus && clampedMinus == minusEnd && Math.Abs(from - middle) > 0.01D)
                AddErrorBarLine(drawing, !horizontal, from, crossPixel - ErrorBarCapHalfLength, crossPixel + ErrorBarCapHalfLength, color, width, OfficeStrokeDashStyle.Solid);
        }
    }

    /// <summary>
    /// Resolves the plotted centre and the plotted ends of one point's error bar. On a logarithmic or reversed axis the
    /// span is measured on the source values, as Office does, and each end is plotted on its own; an end a logarithmic
    /// axis cannot draw, at or below zero, lies past the axis edge and is returned as an infinity.
    /// </summary>
    private static bool TryGetPlottedErrorSpan(OfficeSeries series, OfficeChartErrorBars bars, int index, double value, double offset,
        bool xBars, IReadOnlyList<double>? spanValues, out double center, out double minusEnd, out double plusEnd) {
        center = minusEnd = plusEnd = double.NaN;
        Func<double, double>? toPlot = xBars ? series.PlotXValue : series.PlotValue;
        Func<double, double>? fromPlot = xBars ? series.SourceXValue : series.ValueFromPlot;
        IReadOnlyList<double>? source = xBars ? series.SourcePointXValues : series.SourcePointValues;
        if (toPlot == null || fromPlot == null || source == null) {
            IReadOnlyList<double> values = spanValues ?? (xBars ? series.XValues ?? Array.Empty<double>() : series.Values);
            if (!bars.TryGetSpan(values, index, value, out double c, out double low, out double high)) return false;
            center = c + offset;
            minusEnd = low + offset;
            plusEnd = high + offset;
            return true;
        }

        if (index >= source.Count || !IsFiniteChartValue(source[index])) return false;
        double sourceValue = source[index];
        if (!bars.TryGetSpan(source, index, sourceValue, out double sourceCenter, out double sourceLow, out double sourceHigh)) return false;
        // A stacked point sits at its running total; the bar keeps its source offsets from there.
        double top = offset == 0D ? sourceValue : fromPlot(value + offset);
        if (!IsFiniteChartValue(top)) return false;
        center = toPlot(top + sourceCenter - sourceValue);
        if (!IsFiniteChartValue(center)) return false;
        minusEnd = PlotErrorBarEnd(toPlot, top + sourceLow - sourceValue, center);
        plusEnd = PlotErrorBarEnd(toPlot, top + sourceHigh - sourceValue, center);
        return true;
    }

    // An end the axis cannot plot lies past the edge where its smallest drawable values are.
    private static double PlotErrorBarEnd(Func<double, double> toPlot, double sourceEnd, double plottedCenter) {
        double plotted = toPlot(sourceEnd);
        if (IsFiniteChartValue(plotted)) return plotted;
        double smallest = toPlot(double.Epsilon);
        return smallest < plottedCenter ? double.NegativeInfinity : double.PositiveInfinity;
    }

    // A horizontal line runs along x at the given y; a vertical line runs along y at the given x.
    private static void AddErrorBarLine(OfficeDrawing drawing, bool horizontal, double cross, double start, double end,
        OfficeColor color, double width, OfficeStrokeDashStyle dash) {
        OfficePoint first = horizontal ? new OfficePoint(start, cross) : new OfficePoint(cross, start);
        OfficePoint second = horizontal ? new OfficePoint(end, cross) : new OfficePoint(cross, end);
        AddPointLine(drawing, new[] { first, second }, color, width, dash);
    }
}
