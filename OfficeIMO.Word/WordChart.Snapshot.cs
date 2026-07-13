using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        private const uint MaxCachedChartPoints = 10000U;

        /// <summary>
        /// Tries to create a dependency-free chart snapshot from cached Word chart data.
        /// </summary>
        public bool TryGetSnapshot(out WordChartSnapshot snapshot) {
            try {
                C.Chart? chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null) {
                    snapshot = null!;
                    return false;
                }

                A.ColorScheme? colorScheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;

                if (!HasSingleSupportedChartElement(plotArea)) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    WordChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetBarChartSnapshotKind(barChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Bar3DChart>() is C.Bar3DChart bar3DChart) {
                    WordChartData? data = ReadCategorySeriesData(bar3DChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetBar3DChartSnapshotKind(bar3DChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    WordChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetLineChartSnapshotKind(lineChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Line3DChart>() is C.Line3DChart line3DChart) {
                    WordChartData? data = ReadCategorySeriesData(line3DChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Line, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                    WordChartData? data = ReadCategorySeriesData(areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetAreaChartSnapshotKind(areaChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Area3DChart>() is C.Area3DChart area3DChart) {
                    WordChartData? data = ReadCategorySeriesData(area3DChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetArea3DChartSnapshotKind(area3DChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.RadarChart>() is C.RadarChart radarChart) {
                    WordChartData? data = ReadCategorySeriesData(radarChart.Elements<C.RadarChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Radar, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    WordChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart) {
                    WordChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Pie3DChart>() is C.Pie3DChart pie3DChart) {
                    WordChartData? data = ReadCategorySeriesData(pie3DChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    WordChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Doughnut, data);
                    return true;
                }

                snapshot = null!;
                return false;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static bool HasSingleSupportedChartElement(C.PlotArea plotArea) {
            int chartElementCount = 0;
            int supportedChartElementCount = 0;

            foreach (OpenXmlElement child in plotArea.ChildElements) {
                if (!IsChartPlotElement(child)) {
                    continue;
                }

                chartElementCount++;
                if (IsSupportedChartPlotElement(child)) {
                    supportedChartElementCount++;
                }
            }

            return chartElementCount == 1 && supportedChartElementCount == 1;
        }

        private static bool IsChartPlotElement(OpenXmlElement element) =>
            element != null &&
            element.LocalName != null &&
            element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase);

        private static bool IsSupportedChartPlotElement(OpenXmlElement element) {
            return element is C.BarChart
                || element is C.Bar3DChart
                || element is C.LineChart
                || element is C.Line3DChart
                || element is C.AreaChart
                || element is C.Area3DChart
                || element is C.RadarChart
                || element is C.ScatterChart
                || element is C.PieChart
                || element is C.Pie3DChart
                || element is C.DoughnutChart;
        }

        private WordChartSnapshot CreateSnapshot(C.Chart chart, WordChartSnapshotKind kind, WordChartData data) {
            return new WordChartSnapshot(
                ReadDrawingName(),
                ReadTitle(chart),
                kind,
                data,
                GetWidthPoints(),
                GetHeightPoints());
        }

        private static WordChartSnapshotKind GetBarChartSnapshotKind(C.BarChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            return MapBarKind(direction, grouping);
        }

        private static WordChartSnapshotKind GetBar3DChartSnapshotKind(C.Bar3DChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            return MapBarKind(direction, grouping);
        }

        private static WordChartSnapshotKind MapBarKind(C.BarDirectionValues direction, C.BarGroupingValues grouping) {
            bool horizontal = direction == C.BarDirectionValues.Bar;

            if (grouping == C.BarGroupingValues.Stacked) {
                return horizontal ? WordChartSnapshotKind.StackedBar : WordChartSnapshotKind.StackedColumn;
            }

            if (grouping == C.BarGroupingValues.PercentStacked) {
                return horizontal ? WordChartSnapshotKind.StackedBar100 : WordChartSnapshotKind.StackedColumn100;
            }

            return horizontal ? WordChartSnapshotKind.ClusteredBar : WordChartSnapshotKind.ClusteredColumn;
        }

        private static WordChartSnapshotKind GetLineChartSnapshotKind(C.LineChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedLine;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedLine100;
            }

            return WordChartSnapshotKind.Line;
        }

        private static WordChartSnapshotKind GetAreaChartSnapshotKind(C.AreaChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedArea100;
            }

            return WordChartSnapshotKind.Area;
        }

        private static WordChartSnapshotKind GetArea3DChartSnapshotKind(C.Area3DChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedArea100;
            }

            return WordChartSnapshotKind.Area;
        }

        private static WordChartData? ReadCategorySeriesData(IEnumerable<OpenXmlCompositeElement> seriesElements, A.ColorScheme? colorScheme) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = Array.Empty<string>();
            int fallbackCategoryCount = 0;
            for (int i = 0; i < seriesList.Count; i++) {
                IReadOnlyList<double> values = ReadCachedNumbers(seriesList[i].GetFirstChild<C.Values>());
                if (values.Count == 0) {
                    continue;
                }

                categories = ReadCachedStrings(seriesList[i].GetFirstChild<C.CategoryAxisData>());
                if (categories.Count > 0) {
                    break;
                }

                fallbackCategoryCount = Math.Max(fallbackCategoryCount, values.Count);
            }

            if (categories.Count == 0) {
                categories = CreateFallbackCategories(fallbackCategoryCount);
                if (categories.Count == 0) {
                    return null;
                }
            }

            var series = new List<WordChartSeries>();
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

                series.Add(new WordChartSeries(
                    name,
                    values,
                    color: ReadSeriesColor(seriesElement, colorScheme),
                    pointColors: ReadPointColors(seriesElement, values.Count, colorScheme)));
            }

            return series.Count == 0 ? null : new WordChartData(categories, series);
        }

        private static WordChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements, A.ColorScheme? colorScheme) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            var series = new List<WordChartSeries>();
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

                series.Add(new WordChartSeries(
                    name,
                    values,
                    xValues.Take(pointCount).ToList(),
                    ReadSeriesColor(seriesElement, colorScheme),
                    ReadPointColors(seriesElement, values.Count, colorScheme)));
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            var categories = categoryXValues
                .Select(value => value.ToString(CultureInfo.InvariantCulture))
                .ToList();
            return new WordChartData(categories, series);
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

        private static OfficeIMO.Drawing.OfficeColor? ReadSeriesColor(OpenXmlElement seriesElement, A.ColorScheme? colorScheme) {
            C.ChartShapeProperties? shapeProperties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            return ReadShapeFillColor(shapeProperties, colorScheme)
                ?? ReadShapeOutlineColor(shapeProperties, colorScheme);
        }

        private static IReadOnlyList<OfficeIMO.Drawing.OfficeColor?>? ReadPointColors(OpenXmlElement seriesElement, int pointCount, A.ColorScheme? colorScheme) {
            if (pointCount <= 0) {
                return null;
            }

            var colors = new OfficeIMO.Drawing.OfficeColor?[pointCount];
            bool hasColor = false;
            foreach (C.DataPoint dataPoint in seriesElement.Elements<C.DataPoint>()) {
                int index = (int)(dataPoint.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue);
                if (index < 0 || index >= pointCount) {
                    continue;
                }

                OfficeIMO.Drawing.OfficeColor? color = ReadShapeFillColor(dataPoint.GetFirstChild<C.ChartShapeProperties>(), colorScheme);
                if (color.HasValue) {
                    colors[index] = color.Value;
                    hasColor = true;
                }
            }

            return hasColor ? colors : null;
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadShapeFillColor(C.ChartShapeProperties? shapeProperties, A.ColorScheme? colorScheme) {
            A.SolidFill? fill = shapeProperties?.GetFirstChild<A.SolidFill>();
            return ReadSolidFillColor(fill, colorScheme);
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadShapeOutlineColor(C.ChartShapeProperties? shapeProperties, A.ColorScheme? colorScheme) {
            A.SolidFill? fill = shapeProperties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>();
            return ReadSolidFillColor(fill, colorScheme);
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadSolidFillColor(A.SolidFill? fill, A.ColorScheme? colorScheme) {
            return OfficeOpenXmlThemeColorResolver.ResolveColor(fill, colorScheme);
        }

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<string>();
            }

            List<C.StringPoint> stringPoints = container.Descendants<C.StringPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            List<C.NumericPoint> numericPoints = container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
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

            List<C.NumericPoint> points = container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
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

        private static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex) {
            uint? pointCount = container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value;
            uint maxIndex = 0U;
            bool hasIndexedPoint = false;
            for (int i = 0; i < points.Count; i++) {
                uint? index = getIndex(points[i]);
                if (!index.HasValue) {
                    continue;
                }

                hasIndexedPoint = true;
                if (index.Value > maxIndex) {
                    maxIndex = index.Value;
                }
            }

            uint indexedLength = hasIndexedPoint ? maxIndex + 1U : (uint)points.Count;
            uint length = Math.Max(pointCount ?? 0U, indexedLength);
            if (length > MaxCachedChartPoints) {
                length = Math.Min(indexedLength, MaxCachedChartPoints);
            }

            if (length > int.MaxValue) {
                return points.Count;
            }

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

        private string ReadDrawingName() {
            return _drawing?.Inline?.DocProperties?.Name?.Value ?? string.Empty;
        }

        private double GetWidthPoints() {
            long emu = _drawing?.Inline?.Extent?.Cx?.Value ?? 0L;
            return emu > 0 ? EmuToPoints(emu) : 450D;
        }

        private double GetHeightPoints() {
            long emu = _drawing?.Inline?.Extent?.Cy?.Value ?? 0L;
            return emu > 0 ? EmuToPoints(emu) : 300D;
        }

        private static double EmuToPoints(long emu) {
            return emu * 72D / EnglishMetricUnitsPerInch;
        }
    }
}
