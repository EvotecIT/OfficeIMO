using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        public bool TryGetSnapshot(out PowerPointChartSnapshot snapshot) {
            try {
                ChartPart chartPart = GetChartPart();
                C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>());
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetBarChartSnapshotKind(barChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>());
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetLineChartSnapshotKind(lineChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>());
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>());
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>());
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

        private static PowerPointChartData? ReadCategorySeriesData(IEnumerable<OpenXmlCompositeElement> seriesElements) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = ReadCachedStrings(seriesList[0].GetFirstChild<C.CategoryAxisData>());
            if (categories.Count == 0) {
                IReadOnlyList<double> firstValues = ReadCachedNumbers(seriesList[0].GetFirstChild<C.Values>());
                categories = CreateFallbackCategories(firstValues.Count);
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

                series.Add(new PowerPointChartSeries(name, values));
            }

            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
        }

        private static PowerPointChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<double> firstXValues = ReadCachedNumbers(seriesList[0].GetFirstChild<C.XValues>());
            if (firstXValues.Count == 0) {
                return null;
            }

            var categories = firstXValues
                .Select(value => value.ToString(CultureInfo.InvariantCulture))
                .ToList();
            var series = new List<PowerPointChartSeries>();
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

                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new PowerPointChartSeries(name, values, xValues.Take(pointCount).ToList()));
            }

            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
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

            List<C.StringPoint> stringPoints = container.Descendants<C.StringPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (stringPoints.Count > 0) {
                return stringPoints.Select(point => point.NumericValue?.Text ?? string.Empty).ToList();
            }

            List<C.NumericPoint> numericPoints = container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (numericPoints.Count > 0) {
                return numericPoints.Select(point => point.NumericValue?.Text ?? string.Empty).ToList();
            }

            return Array.Empty<string>();
        }

        private static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<double>();
            }

            var values = new List<double>();
            foreach (C.NumericPoint point in container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U)) {
                string? text = point.NumericValue?.Text;
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    values.Add(value);
                } else {
                    values.Add(0D);
                }
            }

            return values;
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
