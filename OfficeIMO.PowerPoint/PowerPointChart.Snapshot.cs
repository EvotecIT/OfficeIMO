using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private const int MaxChartCachePoints = 100_000;

        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        internal bool TryGetSnapshot(out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshot(null, out snapshot);

        internal bool TryGetSnapshot(A.ColorScheme? colorScheme, out PowerPointChartSnapshot snapshot) {
            try {
                ChartPart chartPart = GetChartPart();
                C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null) {
                    snapshot = null!;
                    return false;
                }

                if (TryCreateMixedChartSnapshot(chart, plotArea, colorScheme, out snapshot)) {
                    return true;
                }

                if (CountSupportedChartElements(plotArea) > 1) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    PowerPointChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    PowerPointChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    PowerPointChartData? data = ReadCategorySeriesData(areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.RadarChart>() is C.RadarChart radarChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(radarChart.Elements<C.RadarChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Radar, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Radar, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Pie, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Doughnut, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Doughnut, data);
                    return true;
                }

                snapshot = null!;
                return false;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static int CountSupportedChartElements(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>().Count()
                + plotArea.Elements<C.LineChart>().Count()
                + plotArea.Elements<C.AreaChart>().Count()
                + plotArea.Elements<C.RadarChart>().Count()
                + plotArea.Elements<C.ScatterChart>().Count()
                + plotArea.Elements<C.PieChart>().Count()
                + plotArea.Elements<C.DoughnutChart>().Count();
        }

        private bool TryCreateMixedChartSnapshot(C.Chart chart, C.PlotArea plotArea, A.ColorScheme? colorScheme, out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            if (CountSupportedChartElements(plotArea) <= 1) {
                return false;
            }

            var parts = new List<(PowerPointChartSnapshotKind Kind, PowerPointChartData Data)>();
            foreach (OpenXmlElement element in plotArea.ChildElements) {
                if (element is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, barChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, lineChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, areaChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme);
                    if (data != null) {
                        parts.Add((PowerPointChartSnapshotKind.Scatter, data));
                    }
                }
            }

            if (parts.Count <= 1) {
                return false;
            }

            if (parts.Any(part => part.Kind == PowerPointChartSnapshotKind.Scatter) &&
                parts.Any(part => part.Kind != PowerPointChartSnapshotKind.Scatter)) {
                return false;
            }

            if (parts.Any(part => IsHorizontalBarKind(part.Kind)) &&
                parts.Any(part => !IsHorizontalBarKind(part.Kind))) {
                return false;
            }

            IReadOnlyList<string> categories = parts[0].Data.Categories;
            var series = new List<PowerPointChartSeries>();
            foreach (var part in parts) {
                foreach (PowerPointChartSeries item in part.Data.Series) {
                    if (item.Values.Count == categories.Count || HasAlignedScatterPoints(item)) {
                        series.Add(item);
                    }
                }
            }

            if (series.Count == 0) {
                return false;
            }

            snapshot = CreateSnapshot(chart, parts[0].Kind, new PowerPointChartData(categories, series));
            return true;
        }

        private static bool HasAlignedScatterPoints(PowerPointChartSeries series) =>
            series.XValues != null &&
            series.XValues.Count == series.Values.Count &&
            series.Values.Count > 0;

        private static bool IsHorizontalBarKind(PowerPointChartSnapshotKind kind) =>
            kind == PowerPointChartSnapshotKind.ClusteredBar ||
            kind == PowerPointChartSnapshotKind.StackedBar ||
            kind == PowerPointChartSnapshotKind.StackedBar100;

        private static OfficeChartAxisGroup GetAxisGroup(C.PlotArea plotArea, OpenXmlCompositeElement chart) {
            HashSet<uint> axisIds = new(chart.Elements<C.AxisId>()
                .Where(axis => axis.Val?.Value != null).Select(axis => axis.Val!.Value));
            return plotArea.Elements<C.ValueAxis>().Any(axis =>
                       axis.AxisId?.Val?.Value != null && axisIds.Contains(axis.AxisId.Val.Value) &&
                       (axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Right ||
                        axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Top))
                ? OfficeChartAxisGroup.Secondary
                : OfficeChartAxisGroup.Primary;
        }

        private PowerPointChartSnapshot CreateSnapshot(C.Chart chart, PowerPointChartSnapshotKind kind, PowerPointChartData data) {
            return new PowerPointChartSnapshot(
                Name ?? string.Empty,
                ReadTitle(chart),
                kind,
                data,
                WidthPoints,
                HeightPoints);
        }

        private static PowerPointChartSnapshotKind GetBarChartSnapshotKind(C.BarChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            bool horizontal = direction == C.BarDirectionValues.Bar;

            if (grouping == C.BarGroupingValues.Stacked) {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar : PowerPointChartSnapshotKind.StackedColumn;
            }

            if (grouping == C.BarGroupingValues.PercentStacked) {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar100 : PowerPointChartSnapshotKind.StackedColumn100;
            }

            return horizontal ? PowerPointChartSnapshotKind.ClusteredBar : PowerPointChartSnapshotKind.ClusteredColumn;
        }

        private static PowerPointChartSnapshotKind GetLineChartSnapshotKind(C.LineChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return PowerPointChartSnapshotKind.StackedLine;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return PowerPointChartSnapshotKind.StackedLine100;
            }

            return PowerPointChartSnapshotKind.Line;
        }

        private static PowerPointChartSnapshotKind GetAreaChartSnapshotKind(C.AreaChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return PowerPointChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return PowerPointChartSnapshotKind.StackedArea100;
            }

            return PowerPointChartSnapshotKind.Area;
        }

        private static PowerPointChartData? ReadCategorySeriesData(IEnumerable<OpenXmlCompositeElement> seriesElements,
            PowerPointChartSnapshotKind? chartKind = null, A.ColorScheme? colorScheme = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = Array.Empty<string>();
            for (int i = 0; i < seriesList.Count; i++) {
                IReadOnlyList<double> values = ReadCachedNumbers(seriesList[i].GetFirstChild<C.Values>());
                if (values.Count == 0) {
                    continue;
                }

                categories = ReadCachedStrings(seriesList[i].GetFirstChild<C.CategoryAxisData>());
                if (categories.Count == 0) {
                    categories = CreateFallbackCategories(values.Count);
                }

                if (categories.Count > 0) {
                    break;
                }
            }

            if (categories.Count == 0) {
                return null;
            }

            var series = new List<PowerPointChartSeries>();
            for (int i = 0; i < seriesList.Count; i++) {
                OpenXmlCompositeElement seriesElement = seriesList[i];
                IReadOnlyList<double> values = NormalizeValues(ReadCachedNumbers(seriesElement.GetFirstChild<C.Values>()), categories.Count);
                if (values.Count == 0) {
                    continue;
                }

                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new PowerPointChartSeries(name, values, null, chartKind,
                    ReadSeriesColor(seriesElement, chartKind, colorScheme), ReadSeriesStrokeWidth(seriesElement),
                    axisGroup) {
                    SourceIndex = seriesElement.GetFirstChild<C.Index>()?.Val?.Value
                });
            }

            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
        }

        private static PowerPointChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements, A.ColorScheme? colorScheme = null) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            var series = new List<PowerPointChartSeries>();
            IReadOnlyList<double>? categoryXValues = null;
            for (int i = 0; i < seriesList.Count; i++) {
                C.ScatterChartSeries seriesElement = seriesList[i];
                IReadOnlyList<double> xValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.XValues>());
                IReadOnlyList<double> yValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.YValues>());
                int pointCount = Math.Min(xValues.Count, yValues.Count);
                if (pointCount == 0) {
                    continue;
                }

                IReadOnlyList<double> values = NormalizeValues(yValues, pointCount);
                if (values.Count == 0) {
                    continue;
                }

                categoryXValues ??= xValues.Take(pointCount).ToList();
                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new PowerPointChartSeries(name, values, xValues.Take(pointCount).ToList(),
                    PowerPointChartSnapshotKind.Scatter,
                    ReadSeriesColor(seriesElement, PowerPointChartSnapshotKind.Scatter, colorScheme),
                    ReadSeriesStrokeWidth(seriesElement)) {
                    SourceIndex = seriesElement.GetFirstChild<C.Index>()?.Val?.Value
                });
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            var categories = categoryXValues
                .Select(value => value.ToString(CultureInfo.InvariantCulture))
                .ToList();
            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
        }

        private static OfficeColor? ReadSeriesColor(OpenXmlCompositeElement seriesElement, PowerPointChartSnapshotKind? chartKind, A.ColorScheme? colorScheme) {
            C.ChartShapeProperties? properties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            if (properties == null) {
                return null;
            }

            OfficeColor? fillColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), colorScheme);
            if (IsFilledChartKind(chartKind)) {
                return fillColor;
            }

            OfficeColor? lineColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), colorScheme);
            if (lineColor.HasValue) {
                return lineColor;
            }

            return fillColor;
        }

        private static bool IsFilledChartKind(PowerPointChartSnapshotKind? chartKind) =>
            chartKind == PowerPointChartSnapshotKind.ClusteredColumn ||
            chartKind == PowerPointChartSnapshotKind.StackedColumn ||
            chartKind == PowerPointChartSnapshotKind.StackedColumn100 ||
            chartKind == PowerPointChartSnapshotKind.ClusteredBar ||
            chartKind == PowerPointChartSnapshotKind.StackedBar ||
            chartKind == PowerPointChartSnapshotKind.StackedBar100 ||
            chartKind == PowerPointChartSnapshotKind.Area ||
            chartKind == PowerPointChartSnapshotKind.StackedArea ||
            chartKind == PowerPointChartSnapshotKind.StackedArea100 ||
            chartKind == PowerPointChartSnapshotKind.Pie ||
            chartKind == PowerPointChartSnapshotKind.Doughnut;

        private static double? ReadSeriesStrokeWidth(OpenXmlCompositeElement seriesElement) {
            C.ChartShapeProperties? properties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            long? widthEmus = properties?.GetFirstChild<A.Outline>()?.Width?.Value;
            return widthEmus.HasValue && widthEmus.Value > 0L
                ? PowerPointUnits.ToPoints(widthEmus.Value)
                : null;
        }

        private static string? ReadTitle(C.Chart chart) {
            C.ChartText? chartText = chart.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return null;
            }

            string text = string.Concat(chartText.Descendants<A.Text>().Select(item => item.Text));
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = ReadCachedStrings(chartText);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        private static string ReadSeriesName(OpenXmlElement seriesElement) {
            C.SeriesText? seriesText = seriesElement.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return string.Empty;
            }

            IReadOnlyList<string> cached = ReadCachedStrings(seriesText);
            if (cached.Count > 0) {
                return cached[0] ?? string.Empty;
            }

            string richText = string.Concat(seriesText.Descendants<A.Text>().Select(item => item.Text));
            return richText.Trim();
        }

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<string>();
            }

            List<C.StringPoint> stringPoints = GetBoundedCachedPoints(container.Descendants<C.StringPoint>());
            stringPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            List<C.NumericPoint> numericPoints = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>());
            numericPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (numericPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    numericPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            return Array.Empty<string>();
        }

        private static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<double>();
            }

            List<C.NumericPoint> points = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>());
            points.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (points.Count == 0) {
                return Array.Empty<double>();
            }

            return CreateIndexedCache(
                container,
                points,
                point => point.Index?.Value,
                point => {
                string? text = point.NumericValue?.Text;
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    return value;
                }

                return 0D;
                },
                0D);
        }

        private static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(
            OpenXmlElement container,
            IReadOnlyList<TPoint> points,
            Func<TPoint, uint?> getIndex,
            Func<TPoint, TValue> getValue,
            TValue defaultValue) {
            int length = GetCachedPointLength(container, points, getIndex);
            var values = Enumerable.Repeat(defaultValue, length).ToArray();
            for (int i = 0; i < points.Count; i++) {
                TPoint point = points[i];
                uint? rawIndex = getIndex(point);
                int index = rawIndex.HasValue && rawIndex.Value <= int.MaxValue
                    ? (int)rawIndex.Value
                    : i;
                if (index >= 0 && index < values.Length) {
                    values[index] = getValue(point);
                }
            }

            return values;
        }

        private static List<TPoint> GetBoundedCachedPoints<TPoint>(IEnumerable<TPoint> points) {
            List<TPoint> boundedPoints = points.Take(MaxChartCachePoints + 1).ToList();
            if (boundedPoints.Count > MaxChartCachePoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {MaxChartCachePoints} points.");
            }

            return boundedPoints;
        }

        private static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex) {
            if (points.Count > MaxChartCachePoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {MaxChartCachePoints} points.");
            }

            uint? pointCount = container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value;
            if (pointCount > MaxChartCachePoints) {
                throw new InvalidDataException($"The chart cache declares more than the supported limit of {MaxChartCachePoints} points.");
            }

            uint maxIndex = 0U;
            bool hasIndexedPoint = false;
            for (int i = 0; i < points.Count; i++) {
                uint? index = getIndex(points[i]);
                if (!index.HasValue) {
                    continue;
                }

                if (index.Value >= MaxChartCachePoints) {
                    throw new InvalidDataException($"The chart cache point index exceeds the supported limit of {MaxChartCachePoints} points.");
                }

                hasIndexedPoint = true;
                if (index.Value > maxIndex) {
                    maxIndex = index.Value;
                }
            }

            uint indexedLength = hasIndexedPoint ? maxIndex + 1U : (uint)points.Count;
            uint length = Math.Max(pointCount ?? 0U, indexedLength);
            return (int)length;
        }

        private static IReadOnlyList<string> CreateFallbackCategories(int count) {
            if (count <= 0) {
                return Array.Empty<string>();
            }

            var categories = new List<string>(count);
            for (int i = 0; i < count; i++) {
                categories.Add("Category " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static IReadOnlyList<double> NormalizeValues(IReadOnlyList<double> values, int count) {
            if (count <= 0 || values.Count == 0) {
                return Array.Empty<double>();
            }

            var normalized = new double[count];
            int take = Math.Min(values.Count, count);
            for (int i = 0; i < take; i++) {
                normalized[i] = values[i];
            }

            return normalized;
        }
    }
}
