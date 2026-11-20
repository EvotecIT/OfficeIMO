namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // A logarithmic or reversed value axis is drawn by plotting each value at log(v) and/or -v on an
    // ordinary linear axis. Ticks, gridlines, bars, lines and crossings then need no special cases;
    // labels restate a plotted value as the source value. Secondary, radar and radial axes stay linear.
    private static OfficeChartSnapshot ApplyValueAxisScales(OfficeChartSnapshot snapshot) {
        OfficeChartLayout layout = snapshot.Layout;
        OfficeChartType kind = snapshot.ChartKind;
        if (IsPieChart(kind) || IsDoughnutChart(kind) || IsRadarChart(kind)) return snapshot;
        bool horizontalBars = IsBarChart(kind);
        bool scatter = IsScatterChart(kind);
        // The axis holding the series values, and for scatter charts the X axis.
        bool valueOnHorizontal = horizontalBars;
        bool transformValues = HasScale(layout, valueOnHorizontal);
        bool transformX = scatter && HasScale(layout, horizontal: true);
        if (!transformValues && !transformX) return snapshot;

        Func<double, double>? toPlot = transformValues ? GetPlottingTransform(layout, valueOnHorizontal) : null;
        Func<double, double>? toSource = transformValues ? GetSourceTransform(layout, valueOnHorizontal) : null;
        Func<double, double>? xToPlot = transformX ? GetPlottingTransform(layout, horizontal: true) : null;
        Func<double, double>? xToSource = transformX ? GetSourceTransform(layout, horizontal: true) : null;
        IReadOnlyList<double> sharedX = GetScatterXValues(snapshot.Data.Categories);

        var series = new List<OfficeSeries>(snapshot.Data.Series.Count);
        foreach (OfficeSeries current in snapshot.Data.Series) {
            bool secondary = current.AxisGroup == OfficeChartAxisGroup.Secondary;
            bool bubble = current.BubbleSizes != null;
            bool values = toPlot != null && !secondary && !(bubble && GetLogBase(layout, valueOnHorizontal).HasValue);
            bool xValues = xToPlot != null && !(bubble && layout.HorizontalAxisLogBase.HasValue);
            if (!values && !xValues) {
                series.Add(current);
                continue;
            }

            IReadOnlyList<double> plotted = values ? current.Values.Select(toPlot!).ToArray() : current.Values;
            IReadOnlyList<double>? plottedX = xValues
                ? (current.XValues ?? sharedX.Take(current.Values.Count).ToArray()).Select(xToPlot!).ToArray()
                : current.XValues;
            series.Add(current.WithPlottedValues(plotted, plottedX, values ? toSource : null, xValues ? xToSource : null));
        }

        IReadOnlyList<double> stackTotals = transformValues && GetLogBase(layout, valueOnHorizontal).HasValue
            ? RestackOnLogarithmicAxis(snapshot, series, toPlot!)
            : Array.Empty<double>();
        OfficeChartLayout plottingLayout = layout;
        if (transformValues) plottingLayout = RestateAxis(plottingLayout, valueOnHorizontal, series.Where(item => item.SourceValue != null)
            .SelectMany(item => item.Values).Concat(stackTotals));
        if (transformX) plottingLayout = RestateAxis(plottingLayout, horizontal: true, series.Where(item => item.SourceXValue != null)
            .SelectMany(item => item.XValues!));
        return new OfficeChartSnapshot(snapshot.Name, snapshot.Title, kind, new OfficeChartData(snapshot.Data.Categories, series),
            snapshot.Width, snapshot.Height, snapshot.Style, plottingLayout, snapshot.BubbleScalePercent, snapshot.BubbleSizeMode,
            snapshot.RadialLayout);
    }

    // Office adds stacked values before it scales them. Each stacked series is replotted as the steps between its
    // scaled running totals, so the renderer's own stacking puts every segment between the totals below and above
    // it, and its labels keep the source values. Stacks start at 1, where a logarithmic axis has its zero, which
    // needs non-negative values whose totals are 0 or at least 1; 100% stacks are not drawn on such an axis.
    // Returns the scaled totals, which set the axis bounds.
    private static IReadOnlyList<double> RestackOnLogarithmicAxis(OfficeChartSnapshot snapshot, List<OfficeSeries> series,
        Func<double, double> toPlot) {
        var totals = new List<double>();
        var runningByKind = new Dictionary<OfficeChartType, double[]>();
        int length = snapshot.Data.Series.Count == 0 ? 0 : snapshot.Data.Series.Max(item => item.Values.Count);
        for (int index = 0; index < series.Count; index++) {
            OfficeSeries plotted = series[index];
            if (plotted.SourceValue == null) continue;
            OfficeChartType kind = GetEffectiveSeriesKind(snapshot, plotted);
            if (IsPercentStackedBarOrColumnChart(kind) || IsPercentStackedLineChart(kind) || IsPercentStackedAreaChart(kind))
                throw new NotSupportedException("A 100% stacked series cannot be drawn on a logarithmic value axis.");
            if (!IsStackedBarOrColumnChart(kind) && !IsStackedLineChart(kind) && !IsStackedAreaChart(kind)) continue;

            IReadOnlyList<double> source = snapshot.Data.Series[index].Values;
            if (!runningByKind.TryGetValue(kind, out double[]? running)) runningByKind[kind] = running = new double[length];
            var steps = new double[source.Count];
            for (int point = 0; point < source.Count; point++) {
                double value = source[point];
                if (!IsFiniteChartValue(value)) {
                    steps[point] = double.NaN;
                    continue;
                }

                if (value < 0D)
                    throw new NotSupportedException("A stacked series with negative values cannot be drawn on a logarithmic value axis.");
                double below = running[point];
                double total = below + value;
                if (total == 0D) continue;
                if (total < 1D)
                    throw new NotSupportedException("A stack whose total is below 1 cannot be drawn on a logarithmic value axis.");
                double top = toPlot(total);
                steps[point] = top - (below > 0D ? toPlot(below) : 0D);
                running[point] = total;
                totals.Add(top);
            }

            series[index] = plotted.WithStackedSteps(steps, source);
        }

        return totals;
    }

    private static bool HasScale(OfficeChartLayout layout, bool horizontal) =>
        GetLogBase(layout, horizontal).HasValue || IsReversedValueAxis(layout, horizontal);

    private static double? GetLogBase(OfficeChartLayout layout, bool horizontal) =>
        horizontal ? layout.HorizontalAxisLogBase : layout.VerticalAxisLogBase;

    private static bool IsReversedValueAxis(OfficeChartLayout layout, bool horizontal) =>
        horizontal ? layout.ReverseHorizontalValueAxis : layout.ReverseVerticalValueAxis;

    private static Func<double, double> GetPlottingTransform(OfficeChartLayout layout, bool horizontal) {
        double? logBase = GetLogBase(layout, horizontal);
        double sign = IsReversedValueAxis(layout, horizontal) ? -1D : 1D;
        if (!logBase.HasValue) return value => sign * value;
        double logOfBase = Math.Log(logBase.Value);
        return value => value > 0D && IsFiniteChartValue(value) ? sign * Math.Log(value) / logOfBase : double.NaN;
    }

    private static Func<double, double> GetSourceTransform(OfficeChartLayout layout, bool horizontal) {
        double? logBase = GetLogBase(layout, horizontal);
        double sign = IsReversedValueAxis(layout, horizontal) ? -1D : 1D;
        // Adding zero turns the -0 of a negated zero into 0, so the label reads "0".
        if (!logBase.HasValue) return value => sign * value + 0D;
        double logBaseValue = logBase.Value;
        // Powers are rounded to 12 significant digits so 10^2 labels as 100, not 99.99999999999997.
        return value => RoundSignificant(Math.Pow(logBaseValue, sign * value));
    }

    private static double RoundSignificant(double value) {
        if (value == 0D || !IsFiniteChartValue(value)) return value;
        double scale = Math.Pow(10D, 11 - Math.Floor(Math.Log10(Math.Abs(value))));
        return Math.Round(value * scale) / scale;
    }

    // Restates explicit bounds and the perpendicular axis's crossing in plotting space. A logarithmic axis
    // spans whole powers of its base with one major tick per power unless the source sets its bounds.
    private static OfficeChartLayout RestateAxis(OfficeChartLayout layout, bool horizontal, IEnumerable<double> plottedValues) {
        Func<double, double> toPlot = GetPlottingTransform(layout, horizontal);
        bool reversed = IsReversedValueAxis(layout, horizontal);
        bool logarithmic = GetLogBase(layout, horizontal).HasValue;
        double? minimum = horizontal ? layout.HorizontalAxisMinimum : layout.VerticalAxisMinimum;
        double? maximum = horizontal ? layout.HorizontalAxisMaximum : layout.VerticalAxisMaximum;
        double? majorUnit = horizontal ? layout.HorizontalAxisMajorUnit : layout.VerticalAxisMajorUnit;
        double? minorUnit = horizontal ? layout.HorizontalAxisMinorUnit : layout.VerticalAxisMinorUnit;
        OfficeChartAxisCrossingPosition crossing = horizontal ? layout.VerticalAxisCrossingPosition : layout.HorizontalAxisCrossingPosition;
        double? crossesAt = horizontal ? layout.VerticalAxisCrossesAt : layout.HorizontalAxisCrossesAt;

        double? plotMinimum = minimum.HasValue ? ToFinite(toPlot(minimum.Value)) : null;
        double? plotMaximum = maximum.HasValue ? ToFinite(toPlot(maximum.Value)) : null;
        if (reversed) (plotMinimum, plotMaximum) = (plotMaximum, plotMinimum);
        if (logarithmic) {
            double[] finite = plottedValues.Where(IsFiniteChartValue).ToArray();
            double low = finite.Length == 0 ? 0D : finite.Min();
            double high = finite.Length == 0 ? 1D : finite.Max();
            plotMinimum ??= Math.Floor(low);
            plotMaximum ??= Math.Ceiling(high);
            if (plotMaximum <= plotMinimum) plotMaximum = plotMinimum + 1D;
            majorUnit = 1D;
            minorUnit = null;
        }

        if (crossesAt.HasValue) crossesAt = ToFinite(toPlot(crossesAt.Value));
        if (reversed && crossing == OfficeChartAxisCrossingPosition.Maximum) crossing = OfficeChartAxisCrossingPosition.Minimum;
        else if (reversed && crossing == OfficeChartAxisCrossingPosition.Minimum) crossing = OfficeChartAxisCrossingPosition.Maximum;
        return layout.WithPlottingScale(horizontal, plotMinimum, plotMaximum, majorUnit, minorUnit, crossing, crossesAt);
    }

    private static double? ToFinite(double value) => IsFiniteChartValue(value) ? value : null;

    // An axis tick restated as the value its label shows.
    private static double GetAxisSourceValue(double plotted, OfficeChartLayout layout, bool horizontal) =>
        HasScale(layout, horizontal) ? GetSourceTransform(layout, horizontal)(plotted) : plotted;

    // Office draws the minor ticks of a logarithmic axis at 2, 3 ... 9 times each power of the base.
    private static IReadOnlyList<double>? GetLogarithmicMinorTicks(OfficeChartLayout layout, bool horizontal, ValueRange range) {
        double? logBase = GetLogBase(layout, horizontal);
        if (!logBase.HasValue) return null;
        bool reversed = IsReversedValueAxis(layout, horizontal);
        double sign = reversed ? -1D : 1D;
        double logOfBase = Math.Log(logBase.Value);
        var ticks = new List<double>();
        int first = (int)Math.Floor(Math.Min(sign * range.Min, sign * range.Max));
        int last = (int)Math.Ceiling(Math.Max(sign * range.Min, sign * range.Max));
        for (int power = first; power < last && ticks.Count < 512; power++) {
            for (int multiple = 2; multiple < logBase.Value && multiple <= 9 && ticks.Count < 512; multiple++) {
                double plotted = sign * (power + Math.Log(multiple) / logOfBase);
                if (plotted > range.Min && plotted < range.Max) ticks.Add(plotted);
            }
        }
        ticks.Sort();
        return ticks;
    }
}
