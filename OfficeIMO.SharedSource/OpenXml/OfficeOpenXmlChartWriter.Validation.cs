using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        internal static void ValidateSharedChartData(OfficeChartData data, OfficeChartKind defaultKind) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (!Enum.IsDefined(typeof(OfficeChartKind), defaultKind))
                throw new ArgumentOutOfRangeException(nameof(defaultKind), "Chart kind is not supported.");
            if (data.Series.Any(series => series.RenderKind.HasValue && !Enum.IsDefined(typeof(OfficeChartKind), series.RenderKind.Value) ||
                !Enum.IsDefined(typeof(OfficeChartAxisGroup), series.AxisGroup)))
                throw new ArgumentException("Series chart kinds and axis groups must be supported values.", nameof(data));
            ValidateSharedWorkbookDimensionsAndValues(data, defaultKind);
            if (defaultKind == OfficeChartKind.Scatter && data.Series.Count > SpreadsheetMaximumColumns / 2)
                throw new ArgumentException("Scatter chart data exceeds the embedded worksheet column limit.", nameof(data));
            if (defaultKind == OfficeChartKind.Bubble) {
                int maximumPoints = data.Series
                    .Select(series => series.Values.Count)
                    .DefaultIfEmpty(0)
                    .Max();
                long totalPoints = data.Series.Sum(series =>
                    (long)series.Values.Count);
                ValidateBubbleWorkbookDimensions(
                    data.Series.Count, maximumPoints, totalPoints);
            }
            for (int index = 0; index < data.Series.Count; index++) {
                OfficeChartSeries series = data.Series[index];
                if (series.Values.Count == 0) {
                    throw new ArgumentException("Chart series cannot be empty.", nameof(data));
                }
                if (defaultKind == OfficeChartKind.Scatter || defaultKind == OfficeChartKind.Bubble) {
                    if (series.XValues != null && series.XValues.Count != series.Values.Count) {
                        throw new ArgumentException("Numeric X and Y value counts must match.", nameof(data));
                    }
                    if (defaultKind == OfficeChartKind.Bubble &&
                        (series.BubbleSizes == null || series.BubbleSizes.Count != series.Values.Count)) {
                        throw new ArgumentException(
                            "Every bubble series must provide one bubble size for each X/Y point.", nameof(data));
                    }
                } else if (series.Values.Count != data.Categories.Count) {
                    throw new ArgumentException("Every chart series must match the category count.", nameof(data));
                }
            }

            List<SharedSeriesDescriptor> descriptors = DescribeSharedSeries(data, defaultKind);
            if (descriptors.Any(item => item.Series.PointExplosions != null &&
                    item.Kind is not OfficeChartKind.Pie and not OfficeChartKind.Doughnut))
                throw new NotSupportedException("Point explosions require a pie or doughnut chart.");
            bool hasSecondary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Secondary);
            if (hasSecondary && descriptors.All(item => item.AxisGroup == OfficeChartAxisGroup.Secondary)) {
                throw new NotSupportedException("A secondary-axis chart requires at least one primary-axis series.");
            }

            if (defaultKind == OfficeChartKind.Scatter || defaultKind == OfficeChartKind.Bubble ||
                descriptors.Any(item => item.Kind == OfficeChartKind.Scatter ||
                                        item.Kind == OfficeChartKind.Bubble)) {
                bool sameNumericKind = descriptors.All(item => item.Kind == defaultKind);
                if ((defaultKind != OfficeChartKind.Scatter && defaultKind != OfficeChartKind.Bubble) ||
                    !sameNumericKind || hasSecondary) {
                    throw new NotSupportedException(
                        "Scatter and bubble series cannot be combined with other chart families or secondary axes.");
                }
                foreach (OfficeChartSeries series in data.Series) {
                    if (series.XValues == null) ParseScatterCategories(data.Categories);
                }
                return;
            }

            bool hasHorizontalBar = descriptors.Any(item => IsHorizontalBarKind(item.Kind));
            if (hasHorizontalBar && (descriptors.Any(item => !IsHorizontalBarKind(item.Kind)) || hasSecondary)) {
                throw new NotSupportedException("Horizontal bar charts cannot be mixed with other families or secondary axes.");
            }

            bool hasStandalone = descriptors.Any(item => item.Kind == OfficeChartKind.Pie ||
                item.Kind == OfficeChartKind.Doughnut || item.Kind == OfficeChartKind.Radar);
            if (hasStandalone && (descriptors.Select(item => item.Kind).Distinct().Count() > 1 || hasSecondary)) {
                throw new NotSupportedException("Pie, doughnut, and radar charts cannot participate in combo or secondary-axis charts.");
            }
        }

        private static void ValidateSharedWorkbookDimensionsAndValues(
            OfficeChartData data, OfficeChartKind defaultKind) {
            const int maximumCellTextLength = 32767;
            long totalPoints = data.Series.Sum(series =>
                (long)series.Values.Count);
            ValidateSharedWorkbookDimensions(data.Categories.Count,
                data.Series.Count, totalPoints);
            if (data.Categories.Any(category => category != null && category.Length > maximumCellTextLength)) {
                throw new ArgumentException("Chart category text exceeds the embedded worksheet cell limit.", nameof(data));
            }
            foreach (OfficeChartSeries series in data.Series) {
                int generatedHeaderLength = defaultKind == OfficeChartKind.Bubble ? 5 :
                    defaultKind == OfficeChartKind.Scatter ? 2 : 0;
                if (series.Name.Length > maximumCellTextLength - generatedHeaderLength) {
                    throw new ArgumentException("Chart series text exceeds the embedded worksheet cell limit, including generated headers.", nameof(data));
                }
                if (series.Values.Any(value => double.IsNaN(value)
                        || double.IsInfinity(value))
                    || series.XValues?.Any(value => double.IsNaN(value)
                        || double.IsInfinity(value)) == true
                    || series.BubbleSizes?.Any(value => double.IsNaN(value)
                        || double.IsInfinity(value)) == true) {
                    throw new ArgumentOutOfRangeException(nameof(data),
                        "Chart data must contain only finite numeric values.");
                }
            }
        }

        internal static void ValidateSharedWorkbookDimensions(
            int categoryCount, int seriesCount, long totalPoints) {
            if (categoryCount > SpreadsheetMaximumRows - 1) {
                throw new ArgumentException(
                    "Chart data exceeds the embedded worksheet row limit.",
                    "data");
            }
            if (seriesCount > SpreadsheetMaximumColumns - 1) {
                throw new ArgumentException(
                    "Chart data exceeds the embedded worksheet column limit.",
                    "data");
            }
            if (totalPoints > MaximumSharedChartPoints) {
                throw new ArgumentException(
                    "Chart data exceeds the shared chart total point limit.",
                    "data");
            }
        }

        internal static void ValidateBubbleWorkbookDimensions(
            int seriesCount, int maximumPoints, long totalPoints) {
            if (seriesCount >
                SpreadsheetMaximumColumns / BubbleWorkbookColumnsPerSeries) {
                throw new ArgumentException(
                    "Bubble chart data exceeds the embedded worksheet column limit.",
                    "data");
            }
            if (maximumPoints > SpreadsheetMaximumRows - 1) {
                throw new ArgumentException(
                    "Bubble chart data exceeds the embedded worksheet row limit.",
                    "data");
            }
            if (maximumPoints > MaximumSharedChartPoints) {
                throw new ArgumentException(
                    "Bubble chart data exceeds the shared chart snapshot point limit.",
                    "data");
            }
            if (totalPoints > MaximumSharedChartPoints) {
                throw new ArgumentException(
                    "Bubble chart data exceeds the shared chart total point limit.",
                    "data");
            }
        }

    }
}
