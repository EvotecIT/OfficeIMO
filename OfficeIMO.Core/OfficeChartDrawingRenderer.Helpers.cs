using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static bool IsBarChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.BarClustered
        || kind == OfficeChartKind.BarStacked
        || kind == OfficeChartKind.BarStacked100;

    private static bool IsColumnChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.ColumnClustered
        || kind == OfficeChartKind.ColumnStacked
        || kind == OfficeChartKind.ColumnStacked100;

    private static bool IsBarOrColumnChart(OfficeChartKind kind) =>
        IsBarChart(kind) || IsColumnChart(kind);

    private static bool IsLineChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.Line
        || kind == OfficeChartKind.LineStacked
        || kind == OfficeChartKind.LineStacked100;

    private static bool IsStackedLineChart(OfficeChartKind kind) => kind == OfficeChartKind.LineStacked;

    private static bool IsPercentStackedLineChart(OfficeChartKind kind) => kind == OfficeChartKind.LineStacked100;

    private static bool IsAreaChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.Area
        || kind == OfficeChartKind.AreaStacked
        || kind == OfficeChartKind.AreaStacked100;

    private static bool IsScatterChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.Scatter || kind == OfficeChartKind.Bubble;

    private static bool IsRadarChart(OfficeChartKind kind) => kind == OfficeChartKind.Radar;

    private static bool IsStackedAreaChart(OfficeChartKind kind) => kind == OfficeChartKind.AreaStacked;

    private static bool IsPercentStackedAreaChart(OfficeChartKind kind) => kind == OfficeChartKind.AreaStacked100;

    private static bool IsStackedBarOrColumnChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.ColumnStacked
        || kind == OfficeChartKind.BarStacked;

    private static bool IsPercentStackedBarOrColumnChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.ColumnStacked100
        || kind == OfficeChartKind.BarStacked100;

    private static bool IsPieChart(OfficeChartKind kind) => kind == OfficeChartKind.Pie;

    private static bool IsDoughnutChart(OfficeChartKind kind) => kind == OfficeChartKind.Doughnut;

    private static OfficeChartKind GetEffectiveSeriesKind(OfficeChartSnapshot snapshot, OfficeChartSeries series) =>
        series.RenderKind ?? snapshot.ChartKind;

    private static bool HasMixedCartesianSeriesKinds(OfficeChartSnapshot snapshot) {
        OfficeChartKind? first = null;
        foreach (OfficeChartSeries series in snapshot.Data.Series) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, series);
            if (IsPieChart(kind) || IsDoughnutChart(kind) || IsRadarChart(kind)) {
                continue;
            }

            if (!first.HasValue) {
                first = kind;
                continue;
            }

            if (first.Value != kind) {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldRenderSeriesAsBarOrColumn(OfficeChartSnapshot snapshot, OfficeChartSeries series) =>
        !HasMixedCartesianSeriesKinds(snapshot) || IsBarOrColumnChart(GetEffectiveSeriesKind(snapshot, series));

    private static bool ShouldRenderSeriesAsArea(OfficeChartSnapshot snapshot, OfficeChartSeries series) =>
        !HasMixedCartesianSeriesKinds(snapshot) || IsAreaChart(GetEffectiveSeriesKind(snapshot, series));

    private static bool ShouldRenderSeriesAsLine(OfficeChartSnapshot snapshot, OfficeChartSeries series) =>
        !HasMixedCartesianSeriesKinds(snapshot) || IsLineChart(GetEffectiveSeriesKind(snapshot, series));

    private static bool ShouldRenderSeriesAsScatter(OfficeChartSnapshot snapshot, OfficeChartSeries series) =>
        !HasMixedCartesianSeriesKinds(snapshot) || IsScatterChart(GetEffectiveSeriesKind(snapshot, series));

    private static double GetPositiveCategoryTotal(IReadOnlyList<OfficeChartSeries> series, int categoryIndex) {
        double total = 0D;
        for (int s = 0; s < series.Count; s++) {
            if (TryGetSeriesValue(series[s], categoryIndex, out double value)) {
                total += Math.Max(0D, value);
            }
        }

        return total;
    }

    private static double GetDataLabelCategoryTotal(IReadOnlyList<OfficeChartSeries> series, int categoryIndex) {
        double total = GetPositiveCategoryTotal(series, categoryIndex);
        if (total > 0D) {
            return total;
        }

        for (int s = 0; s < series.Count; s++) {
            if (TryGetSeriesValue(series[s], categoryIndex, out double value)) {
                total += Math.Abs(value);
            }
        }

        return total;
    }

    private readonly struct PercentStackedTotals {
        internal PercentStackedTotals(double[] positive, double[] negative) {
            Positive = positive;
            Negative = negative;
        }

        internal double[] Positive { get; }
        internal double[] Negative { get; }
    }

    private static PercentStackedTotals BuildPercentStackedTotals(IReadOnlyList<OfficeChartSeries> series, int categoryCount) {
        var positive = new double[categoryCount];
        var negative = new double[categoryCount];
        for (int s = 0; s < series.Count; s++) {
            for (int categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
                if (!TryGetSeriesValue(series[s], categoryIndex, out double value)) {
                    continue;
                }

                if (value > 0D) {
                    positive[categoryIndex] += value;
                } else if (value < 0D) {
                    negative[categoryIndex] += Math.Abs(value);
                }
            }
        }

        return new PercentStackedTotals(positive, negative);
    }

    private static double NormalizePercentStackedValue(PercentStackedTotals totals, int categoryIndex, double value) {
        if (value == 0D) {
            return 0D;
        }

        double total = value > 0D ? totals.Positive[categoryIndex] : totals.Negative[categoryIndex];
        return total <= 0D ? 0D : value / total;
    }

    private static ValueRange GetPercentStackedSeriesRange(IReadOnlyList<OfficeChartSeries> series, int categoryCount) {
        PercentStackedTotals totals = BuildPercentStackedTotals(series, categoryCount);
        for (int category = 0; category < categoryCount; category++) {
            if (totals.Negative[category] > 0D) {
                return new ValueRange(-1D, 1D);
            }
        }

        return new ValueRange(0D, 1D);
    }

    private static ValueRange GetStackedSeriesRange(IReadOnlyList<OfficeChartSeries> series, int categoryCount) {
        double min = 0D;
        double max = 0D;
        for (int category = 0; category < categoryCount; category++) {
            double positive = 0D;
            double negative = 0D;
            for (int s = 0; s < series.Count; s++) {
                if (!TryGetSeriesValue(series[s], category, out double value)) {
                    continue;
                }

                if (value >= 0D) {
                    positive += value;
                } else {
                    negative += value;
                }
            }

            if (positive > max) {
                max = positive;
            }

            if (negative < min) {
                min = negative;
            }
        }

        return ExpandFlatRange(min, max);
    }

    private static bool TryGetSeriesValue(OfficeChartSeries series, int index, out double value) {
        value = 0D;
        if (index < 0 || index >= series.Values.Count) {
            return false;
        }

        value = series.Values[index];
        return IsFiniteChartValue(value);
    }

    private static double GetSeriesValue(OfficeChartSeries series, int index) {
        return TryGetSeriesValue(series, index, out double value) ? value : 0D;
    }

    private static bool IsFiniteChartValue(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static double ToPlotY(double value, double min, double max, double plotTop, double plotHeight) {
        double ratio = GetPlotRatio(value, min, max);
        return plotTop + plotHeight - plotHeight * ratio;
    }

    private static double ToPlotX(double value, double min, double max, double plotLeft, double plotWidth) {
        double ratio = GetPlotRatio(value, min, max);
        return plotLeft + plotWidth * ratio;
    }

    private static double GetPlotRatio(double value, double min, double max) {
        double ratio = GetUnclampedPlotRatio(value, min, max);
        if (ratio < 0D) {
            ratio = 0D;
        } else if (ratio > 1D) {
            ratio = 1D;
        }
        return ratio;
    }

    private static double GetUnclampedPlotRatio(double value, double min, double max) {
        double range = max - min;
        return range <= 0D ? 0.5D : double.IsInfinity(range)
            ? (value * 0.5D - min * 0.5D) / (max * 0.5D - min * 0.5D)
            : (value - min) / range;
    }

    private static IReadOnlyList<double> GetScatterXValues(IReadOnlyList<string> categories) {
        var values = new double[categories.Count];
        for (int i = 0; i < categories.Count; i++) {
            if (double.TryParse(categories[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                !double.IsNaN(value) &&
                !double.IsInfinity(value)) {
                values[i] = value;
            } else {
                values[i] = i + 1D;
            }
        }

        return values;
    }

    private static ValueRange GetScatterXRange(IReadOnlyList<OfficeChartSeries> series, IReadOnlyList<double> sharedXValues) {
        var values = new List<double>(sharedXValues.Count);
        bool hasSeriesXValues = false;
        bool usesSharedXValues = false;
        for (int s = 0; s < series.Count; s++) {
            IReadOnlyList<double>? xValues = series[s].XValues;
            if (xValues == null) {
                usesSharedXValues = true;
                continue;
            }

            hasSeriesXValues = true;
            values.AddRange(xValues);
        }

        if (hasSeriesXValues && usesSharedXValues) {
            values.AddRange(sharedXValues);
        }

        return hasSeriesXValues ? GetFiniteRange(values) : GetFiniteRange(sharedXValues);
    }

    private static (ValueRange XRange, ValueRange YRange) GetScatterPointRanges(IReadOnlyList<OfficeChartSeries> series, IReadOnlyList<double> sharedXValues) {
        var xValues = new List<double>();
        var yValues = new List<double>();
        for (int s = 0; s < series.Count; s++) {
            IReadOnlyList<double> seriesXValues = series[s].XValues ?? sharedXValues;
            int pointCount = Math.Min(seriesXValues.Count, series[s].Values.Count);
            for (int i = 0; i < pointCount; i++) {
                if (!TryGetSeriesValue(series[s], i, out double yValue)) {
                    continue;
                }

                double xValue = seriesXValues[i];
                if (!IsFiniteChartValue(xValue)) {
                    continue;
                }

                xValues.Add(xValue);
                yValues.Add(yValue);
            }
        }

        return xValues.Count == 0 || yValues.Count == 0
            ? (GetScatterXRange(series, sharedXValues), GetFiniteSeriesRange(series))
            : (GetFiniteRange(xValues), GetFiniteRange(yValues));
    }

    private static IReadOnlyList<OfficePoint> CreateRadarPoints(int count, double centerX, double centerY, double radius) {
        var points = new List<OfficePoint>(count);
        for (int i = 0; i < count; i++) {
            points.Add(CreateRadarPoint(i, count, centerX, centerY, radius));
        }

        return points;
    }

    private static OfficePoint CreateRadarPoint(int index, int count, double centerX, double centerY, double radius) {
        double angle = -Math.PI / 2D + Math.PI * 2D * index / count;
        return new OfficePoint(centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
    }

    private static ValueRange GetFiniteSeriesRange(IReadOnlyList<OfficeChartSeries> series) {
        bool any = false;
        double min = 0D;
        double max = 0D;
        foreach (OfficeChartSeries item in series) {
            foreach (double value in item.Values) {
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    continue;
                }

                if (!any) {
                    min = value;
                    max = value;
                    any = true;
                } else {
                    if (value < min) {
                        min = value;
                    }

                    if (value > max) {
                        max = value;
                    }
                }
            }
        }

        return any ? ExpandFlatRange(min, max) : new ValueRange(0D, 1D);
    }

    private static ValueRange ApplyValueAxisScale(ValueRange range, OfficeChartLayout layout, bool horizontal) {
        double min = horizontal
            ? layout.HorizontalAxisMinimum ?? range.Min
            : layout.VerticalAxisMinimum ?? range.Min;
        double max = horizontal
            ? layout.HorizontalAxisMaximum ?? range.Max
            : layout.VerticalAxisMaximum ?? range.Max;
        if (max <= min) {
            double span = range.Max - range.Min;
            if (double.IsNaN(span) || double.IsInfinity(span) || span <= 0D)
                span = 1D;
            bool explicitMin = horizontal ? layout.HorizontalAxisMinimum.HasValue : layout.VerticalAxisMinimum.HasValue;
            bool explicitMax = horizontal ? layout.HorizontalAxisMaximum.HasValue : layout.VerticalAxisMaximum.HasValue;
            if (!explicitMin && !explicitMax) return range;
            if (explicitMin && !explicitMax) max = min + span;
            else if (explicitMax && !explicitMin) min = max - span;
            else throw new ArgumentException("The value-axis maximum must exceed its minimum.", nameof(layout));
            if (double.IsInfinity(min) || double.IsInfinity(max) || max <= min)
                throw new ArgumentException("The value-axis bounds cannot form a finite range.", nameof(layout));
        }

        return new ValueRange(min, max);
    }

    private static bool HasValueAxisScale(OfficeChartLayout layout, bool horizontal) =>
        horizontal
            ? layout.HorizontalAxisMinimum.HasValue || layout.HorizontalAxisMaximum.HasValue || layout.HorizontalAxisMajorUnit.HasValue || layout.HorizontalAxisMinorUnit.HasValue
            : layout.VerticalAxisMinimum.HasValue || layout.VerticalAxisMaximum.HasValue || layout.VerticalAxisMajorUnit.HasValue || layout.VerticalAxisMinorUnit.HasValue;

    private static double? GetValueAxisMajorUnit(OfficeChartLayout layout, bool horizontal) =>
        horizontal ? layout.HorizontalAxisMajorUnit : layout.VerticalAxisMajorUnit;

    private static double? GetValueAxisMinorUnit(OfficeChartLayout layout, bool horizontal) =>
        horizontal ? layout.HorizontalAxisMinorUnit : layout.VerticalAxisMinorUnit;

    internal static bool HasUnsupportedAxisUnitBudget(OfficeChartSnapshot snapshot) {
        if (IsRadarChart(snapshot.ChartKind))
            return snapshot.Layout.HorizontalAxisMajorUnit.HasValue || snapshot.Layout.HorizontalAxisMinorUnit.HasValue ||
                snapshot.Layout.VerticalAxisMajorUnit.HasValue || snapshot.Layout.VerticalAxisMinorUnit.HasValue;
        if (IsPieChart(snapshot.ChartKind) || IsDoughnutChart(snapshot.ChartKind))
            return false;
        OfficeChartLayout layout = snapshot.Layout;
        bool barChart = IsBarChart(snapshot.ChartKind);
        bool hasSecondary = snapshot.Data.Series.Any(series => series.AxisGroup == OfficeChartAxisGroup.Secondary);
        ValueRange primary = GetPrimaryValueAxisRange(snapshot, layout, barChart, hasSecondary);
        if (ExceedsAxisUnitBudget(primary, GetValueAxisMajorUnit(layout, barChart), 32) ||
            ExceedsMinorAxisUnitBudget(primary, GetValueAxisMajorUnit(layout, barChart),
                GetValueAxisMinorUnit(layout, barChart), barChart
                    ? layout.HorizontalAxisMinorTickMark != OfficeChartAxisTickMark.None
                    : layout.VerticalAxisMinorTickMark != OfficeChartAxisTickMark.None)) return true;
        if (hasSecondary && layout.SecondaryValueAxis is OfficeChartValueAxisLayout secondary) {
            var secondaryLayout = new OfficeChartLayout(
                horizontalAxisMinimum: secondary.Minimum, horizontalAxisMaximum: secondary.Maximum,
                verticalAxisMinimum: secondary.Minimum, verticalAxisMaximum: secondary.Maximum);
            ValueRange secondaryRange = ApplyValueAxisScale(
                GetMixedCartesianValueRange(snapshot, OfficeChartAxisGroup.Secondary),
                secondaryLayout, horizontal: barChart);
            if (ExceedsAxisUnitBudget(secondaryRange, secondary.MajorUnit, 32) ||
                ExceedsMinorAxisUnitBudget(secondaryRange, secondary.MajorUnit,
                    secondary.MinorUnit, (secondary.MinorTickMark ?? layout.VerticalAxisMinorTickMark) !=
                        OfficeChartAxisTickMark.None || (secondary.MinorTickMark ?? layout.HorizontalAxisMinorTickMark) !=
                        OfficeChartAxisTickMark.None)) return true;
        }
        if (!IsScatterChart(snapshot.ChartKind)) return false;
        IReadOnlyList<double> sharedX = GetScatterXValues(snapshot.Data.Categories);
        List<OfficeChartSeries> series = GetRenderableScatterSeries(snapshot).Select(item => item.Series).ToList();
        ValueRange horizontal = ApplyValueAxisScale(GetScatterPointRanges(series, sharedX).XRange,
            layout, horizontal: true);
        return ExceedsAxisUnitBudget(horizontal, layout.HorizontalAxisMajorUnit, 32) ||
            ExceedsAxisUnitBudget(horizontal, layout.HorizontalAxisMinorUnit, 96);
    }

    private static bool ExceedsAxisUnitBudget(ValueRange range, double? unit, int maximumTicks) {
        if (!unit.HasValue || unit.Value <= 0D) return false;
        double intervals = (range.Max - range.Min) / unit.Value;
        return double.IsNaN(intervals) || double.IsInfinity(intervals) ||
            intervals < 1D || intervals >= maximumTicks;
    }

    private static bool ExceedsMinorAxisUnitBudget(ValueRange range, double? majorUnit,
        double? minorUnit, bool automaticWhenVisible) {
        if (!minorUnit.HasValue && automaticWhenVisible)
            minorUnit = (majorUnit ?? (range.Max - range.Min) / 4D) / 5D;
        return ExceedsAxisUnitBudget(range, minorUnit, 96);
    }

    private static IReadOnlyList<double> GetValueAxisMajorTicks(ValueRange range, double? majorUnit) {
        if (!majorUnit.HasValue || majorUnit.Value <= 0D) {
            return new[] {
                range.Min,
                range.Min * 0.75D + range.Max * 0.25D,
                range.Min * 0.5D + range.Max * 0.5D,
                range.Min * 0.25D + range.Max * 0.75D,
                range.Max
            };
        }

        double span = range.Max - range.Min;
        if (span <= 0D) {
            return new[] { range.Min, range.Max };
        }

        if (ExceedsAxisUnitBudget(range, majorUnit, 32))
            throw new NotSupportedException("The chart major-axis unit exceeds the renderer tick budget.");

        int tickCount = (int)Math.Floor(span / majorUnit.Value) + 1;

        var ticks = new List<double>(tickCount + 1);
        for (int i = 0; i < tickCount; i++) {
            double value = range.Min + majorUnit.Value * i;
            if (value > range.Max) {
                break;
            }

            ticks.Add(value);
        }

        if (ticks.Count == 0 || Math.Abs(ticks[ticks.Count - 1] - range.Max) > Math.Max(0.000001D, Math.Abs(range.Max) * 0.0000001D)) {
            ticks.Add(range.Max);
        }

        return ticks;
    }

    private static IReadOnlyList<double> GetValueAxisLabelTicks(ValueRange range, double? majorUnit) =>
        GetValueAxisMajorTicks(range, majorUnit);

    private static IReadOnlyList<double> GetValueAxisMinorTicks(ValueRange range, double? minorUnit,
        IReadOnlyList<double> majorTicks, bool automaticWhenVisible = false) {
        if (!minorUnit.HasValue && automaticWhenVisible && majorTicks.Count >= 2) {
            double interval = majorTicks[1] - majorTicks[0];
            if (interval > 0D && !double.IsInfinity(interval)) minorUnit = interval / 5D;
        }
        if (!minorUnit.HasValue || minorUnit.Value <= 0D) {
            return Array.Empty<double>();
        }

        double span = range.Max - range.Min;
        if (span <= 0D) {
            return Array.Empty<double>();
        }

        if (ExceedsAxisUnitBudget(range, minorUnit, 96))
            throw new NotSupportedException("The chart minor-axis unit exceeds the renderer tick budget.");

        int tickCount = (int)Math.Floor(span / minorUnit.Value) + 1;

        var ticks = new List<double>(tickCount);
        for (int i = 0; i < tickCount; i++) {
            double value = range.Min + minorUnit.Value * i;
            if (value <= range.Min || value >= range.Max) {
                continue;
            }

            if (IsMajorTick(value, majorTicks)) {
                continue;
            }

            ticks.Add(value);
        }

        return ticks;
    }

    private static bool IsMajorTick(double value, IReadOnlyList<double> majorTicks) {
        for (int i = 0; i < majorTicks.Count; i++) {
            if (Math.Abs(value - majorTicks[i]) <= Math.Max(0.000001D, Math.Abs(majorTicks[i]) * 0.0000001D)) {
                return true;
            }
        }

        return false;
    }

    private static int GetCategorySlotIndex(int categoryIndex, int categoryCount, OfficeChartLayout layout) =>
        layout.ReverseCategoryAxis ? categoryCount - 1 - categoryIndex : categoryIndex;

    private static int GetHorizontalBarCategorySlotIndex(int categoryIndex, int categoryCount, OfficeChartLayout layout) =>
        layout.CategoryAxisOrientationSpecified
            ? GetCategorySlotIndex(categoryIndex, categoryCount, layout)
            : categoryCount - 1 - categoryIndex;

    private static double GetCategorySlotCenterX(double plotLeft, double slotWidth, int categoryIndex, int categoryCount, OfficeChartLayout layout) =>
        plotLeft + slotWidth * GetCategorySlotIndex(categoryIndex, categoryCount, layout) + slotWidth / 2D;

    private static double GetCategoryPointX(double plotLeft, double step, int categoryIndex, int categoryCount, OfficeChartLayout layout) =>
        plotLeft + step * GetCategorySlotIndex(categoryIndex, categoryCount, layout);

    private static double GetCartesianCategoryPointX(OfficeChartSnapshot snapshot, double plotLeft, double plotWidth,
        int categoryIndex, int categoryCount, OfficeChartLayout layout) {
        // A column plot uses category slots. Its line and area companions must
        // pass through the same slot centers, including on a reversed axis.
        if (snapshot.Data.Series.Any(series => IsColumnChart(GetEffectiveSeriesKind(snapshot, series)))) {
            return GetCategorySlotCenterX(plotLeft, plotWidth / categoryCount, categoryIndex, categoryCount, layout);
        }

        double step = categoryCount > 1 ? plotWidth / (categoryCount - 1) : 0D;
        return GetCategoryPointX(plotLeft, step, categoryIndex, categoryCount, layout);
    }

    private static ValueRange GetFiniteRange(IReadOnlyList<double> values) {
        bool any = false;
        double min = 0D;
        double max = 0D;
        foreach (double value in values) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                continue;
            }

            if (!any) {
                min = value;
                max = value;
                any = true;
            } else {
                if (value < min) {
                    min = value;
                }

                if (value > max) {
                    max = value;
                }
            }
        }

        return any ? ExpandFlatRange(min, max) : new ValueRange(0D, 1D);
    }

    private static ValueRange GetRadarValueRange(IReadOnlyList<OfficeChartSeries> series) {
        ValueRange range = GetFiniteSeriesRange(series);
        return ExpandFlatRange(Math.Min(0D, range.Min), Math.Max(0D, range.Max));
    }

    private static double ToRadarRadiusRatio(double value, double min, double max) {
        double range = max - min;
        double ratio = range <= 0D ? 0.5D : (value - min) / range;
        if (ratio < 0D) {
            return 0D;
        }

        if (ratio > 1D) {
            return 1D;
        }

        return ratio;
    }

    private static ValueRange ExpandFlatRange(double min, double max) {
        if (max > min) {
            return new ValueRange(min, max);
        }

        double padding = Math.Abs(min) > 1D ? Math.Abs(min) * 0.1D : 1D;
        return new ValueRange(min - padding, max + padding);
    }

    private readonly struct ValueRange {
        public ValueRange(double min, double max) {
            Min = min;
            Max = max;
        }

        public double Min { get; }
        public double Max { get; }
    }
}
