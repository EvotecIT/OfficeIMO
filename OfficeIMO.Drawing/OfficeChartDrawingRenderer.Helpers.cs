using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static bool IsBarChart(OfficeChartKind kind) =>
        kind == OfficeChartKind.BarClustered
        || kind == OfficeChartKind.BarStacked
        || kind == OfficeChartKind.BarStacked100;

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

    private static bool IsScatterChart(OfficeChartKind kind) => kind == OfficeChartKind.Scatter;

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

    private static double GetPercentStackedCategoryTotal(IReadOnlyList<OfficeChartSeries> series, int categoryIndex, bool positive) {
        double total = 0D;
        for (int s = 0; s < series.Count; s++) {
            if (!TryGetSeriesValue(series[s], categoryIndex, out double value)) {
                continue;
            }

            if (positive && value > 0D) {
                total += value;
            } else if (!positive && value < 0D) {
                total += Math.Abs(value);
            }
        }

        return total;
    }

    private static double NormalizePercentStackedValue(IReadOnlyList<OfficeChartSeries> series, int categoryIndex, double value) {
        if (value == 0D) {
            return 0D;
        }

        double total = GetPercentStackedCategoryTotal(series, categoryIndex, value > 0D);
        return total <= 0D ? 0D : value / total;
    }

    private static ValueRange GetPercentStackedSeriesRange(IReadOnlyList<OfficeChartSeries> series, int categoryCount) {
        bool hasNegative = false;
        for (int category = 0; category < categoryCount; category++) {
            if (GetPercentStackedCategoryTotal(series, category, positive: false) > 0D) {
                hasNegative = true;
                break;
            }
        }

        return hasNegative ? new ValueRange(-1D, 1D) : new ValueRange(0D, 1D);
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
        double range = max - min;
        double ratio = range <= 0D ? 0.5D : (value - min) / range;
        if (ratio < 0D) {
            ratio = 0D;
        } else if (ratio > 1D) {
            ratio = 1D;
        }

        return plotTop + plotHeight - plotHeight * ratio;
    }

    private static double ToPlotX(double value, double min, double max, double plotLeft, double plotWidth) {
        double range = max - min;
        double ratio = range <= 0D ? 0.5D : (value - min) / range;
        if (ratio < 0D) {
            ratio = 0D;
        } else if (ratio > 1D) {
            ratio = 1D;
        }

        return plotLeft + plotWidth * ratio;
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
