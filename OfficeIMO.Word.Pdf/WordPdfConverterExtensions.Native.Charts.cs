using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double NativeEmusPerPoint = 12700D;

        private static bool RenderNativeChart(INativePdfFlow pdf, WordChart? chart, PdfCore.PdfAlign align, PdfSaveOptions? options, string source) {
            if (chart == null) {
                return false;
            }

            if (!TryCreateNativeWordChartSnapshot(chart, out OfficeChartSnapshot? snapshot, out string? warning)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeBodyChartUnsupported",
                        source,
                        warning ?? "Word chart data is not mapped by the OfficeIMO PDF engine yet.");
                }

                return false;
            }

            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(snapshot!);
            if (rendering.QualityReport.HasIssues && options != null) {
                AddNativeExportWarning(
                    options,
                    "NativeBodyChartQuality",
                    source,
                    "Exported Word chart '" + GetNativeWordChartDisplayName(snapshot!) + "' with shared drawing quality warnings: " + string.Join("; ", rendering.QualityReport.Issues.Select(issue => issue.ToString())));
            }

            pdf.Drawing(rendering.Drawing, align, spacingBefore: 2D, spacingAfter: 6D);
            return true;
        }

        private static bool TryCreateNativeWordChartSnapshot(WordChart chart, out OfficeChartSnapshot? snapshot, out string? warning) {
            snapshot = null;
            warning = null;

            ChartPart? chartPart = chart.ChartPart;
            Chart? openXmlChart = chartPart?.ChartSpace?.GetFirstChild<Chart>();
            PlotArea? plotArea = openXmlChart?.PlotArea;
            if (plotArea == null) {
                warning = "Word chart part does not contain a plot area with cached chart data.";
                return false;
            }

            List<OpenXmlElement> chartElements = plotArea.ChildElements
                .Where(IsNativeSupportedWordChartElement)
                .ToList();
            if (chartElements.Count == 0) {
                warning = "Word chart type is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            if (chartElements.Count > 1) {
                warning = "Word combo charts with multiple plot types are not supported by the shared OfficeIMO chart renderer yet.";
                return false;
            }

            OpenXmlElement chartElement = chartElements[0];
            if (!TryMapNativeWordChartKind(chartElement, out OfficeChartKind chartKind)) {
                warning = "Word chart type '" + chartElement.LocalName + "' is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            IReadOnlyList<OfficeChartSeries> series = ExtractNativeWordChartSeries(chartElement, chartKind, out IReadOnlyList<string> categories);
            if (categories.Count == 0 || series.Count == 0) {
                warning = "Word chart does not contain cached categories and values that can be rendered without Office.";
                return false;
            }

            (double width, double height) = GetNativeWordChartSizePoints(chart);
            string? title = GetNativeWordChartTitle(openXmlChart!);
            string name = GetNativeWordChartName(chart, chartPart, title);
            OfficeChartLayout? layout = CreateNativeWordChartLayout(chartElement);
            snapshot = new OfficeChartSnapshot(
                name,
                title,
                chartKind,
                new OfficeChartData(categories, series),
                width,
                height,
                layout: layout);
            return true;
        }

        private static bool IsNativeSupportedWordChartElement(OpenXmlElement element) =>
            element.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or "areaChart" or "area3DChart" or "pieChart" or "pie3DChart" or "doughnutChart" or "scatterChart" or "radarChart";

        private static bool TryMapNativeWordChartKind(OpenXmlElement chartElement, out OfficeChartKind kind) {
            switch (chartElement.LocalName) {
                case "barChart":
                case "bar3DChart":
                    bool horizontal = chartElement.GetFirstChild<BarDirection>()?.Val?.Value == BarDirectionValues.Bar;
                    string grouping = chartElement.GetFirstChild<BarGrouping>()?.Val?.Value.ToString() ?? string.Empty;
                    kind = MapNativeWordBarChartKind(horizontal, grouping);
                    return true;
                case "lineChart":
                case "line3DChart":
                    kind = MapNativeWordLineChartKind(chartElement.GetFirstChild<Grouping>()?.Val?.Value.ToString() ?? string.Empty);
                    return true;
                case "areaChart":
                case "area3DChart":
                    kind = MapNativeWordAreaChartKind(chartElement.GetFirstChild<Grouping>()?.Val?.Value.ToString() ?? string.Empty);
                    return true;
                case "pieChart":
                case "pie3DChart":
                    kind = OfficeChartKind.Pie;
                    return true;
                case "doughnutChart":
                    kind = OfficeChartKind.Doughnut;
                    return true;
                case "scatterChart":
                    kind = OfficeChartKind.Scatter;
                    return true;
                case "radarChart":
                    kind = OfficeChartKind.Radar;
                    return true;
                default:
                    kind = default;
                    return false;
            }
        }

        private static OfficeChartKind MapNativeWordBarChartKind(bool horizontal, string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return horizontal ? OfficeChartKind.BarStacked : OfficeChartKind.ColumnStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return horizontal ? OfficeChartKind.BarStacked100 : OfficeChartKind.ColumnStacked100;
            }

            return horizontal ? OfficeChartKind.BarClustered : OfficeChartKind.ColumnClustered;
        }

        private static OfficeChartKind MapNativeWordLineChartKind(string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartKind.LineStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartKind.LineStacked100;
            }

            return OfficeChartKind.Line;
        }

        private static OfficeChartKind MapNativeWordAreaChartKind(string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartKind.AreaStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartKind.AreaStacked100;
            }

            return OfficeChartKind.Area;
        }

        private static IReadOnlyList<OfficeChartSeries> ExtractNativeWordChartSeries(OpenXmlElement chartElement, OfficeChartKind chartKind, out IReadOnlyList<string> categories) {
            var series = new List<OfficeChartSeries>();
            var categoryList = new List<string>();
            bool isScatter = chartKind == OfficeChartKind.Scatter;

            int seriesIndex = 0;
            foreach (OpenXmlElement seriesElement in chartElement.ChildElements.Where(element => element.LocalName == "ser")) {
                IReadOnlyList<double> values;
                IReadOnlyList<double>? xValues = null;
                IReadOnlyList<string> currentCategories;
                if (isScatter) {
                    xValues = ExtractNativeWordChartNumberValues(seriesElement.Elements<XValues>().FirstOrDefault());
                    values = ExtractNativeWordChartNumberValues(seriesElement.Elements<YValues>().FirstOrDefault());
                    if (xValues.Count != values.Count) {
                        int count = Math.Max(xValues.Count, values.Count);
                        xValues = NormalizeNativeWordChartNumberValues(xValues, count, useIndexDefaults: true);
                        values = NormalizeNativeWordChartNumberValues(values, count, useIndexDefaults: false);
                    }

                    currentCategories = xValues.Count > 0
                        ? xValues.Select(value => value.ToString("0.####", CultureInfo.InvariantCulture)).ToList()
                        : CreateNativeWordChartDefaultCategories(values.Count);
                } else {
                    values = ExtractNativeWordChartNumberValues(seriesElement.Elements<Values>().FirstOrDefault());
                    currentCategories = ExtractNativeWordChartCategories(seriesElement.Elements<CategoryAxisData>().FirstOrDefault(), values.Count);
                }

                if (values.Count == 0) {
                    continue;
                }

                if (categoryList.Count == 0) {
                    categoryList.AddRange(currentCategories.Count > 0 ? currentCategories : CreateNativeWordChartDefaultCategories(values.Count));
                }

                series.Add(new OfficeChartSeries(GetNativeWordChartSeriesName(seriesElement, seriesIndex), values, xValues));
                seriesIndex++;
            }

            categories = categoryList;
            return series;
        }

        private static IReadOnlyList<string> ExtractNativeWordChartCategories(OpenXmlElement? categoryAxisData, int valueCount) {
            List<string> categories = ExtractNativeWordChartStringValues(categoryAxisData).ToList();
            if (categories.Count == 0) {
                categories.AddRange(ExtractNativeWordChartNumberValues(categoryAxisData).Select(value => value.ToString("0.####", CultureInfo.InvariantCulture)));
            }

            if (categories.Count == 0 && valueCount > 0) {
                categories.AddRange(CreateNativeWordChartDefaultCategories(valueCount));
            }

            for (int index = 0; index < categories.Count; index++) {
                if (string.IsNullOrWhiteSpace(categories[index])) {
                    categories[index] = "Category " + (index + 1).ToString(CultureInfo.InvariantCulture);
                }
            }

            return categories;
        }

        private static IReadOnlyList<string> CreateNativeWordChartDefaultCategories(int count) {
            var categories = new List<string>();
            for (int i = 0; i < count; i++) {
                categories.Add("Category " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static IReadOnlyList<string> ExtractNativeWordChartStringValues(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<string>();
            }

            var values = new SortedDictionary<uint, string>();
            uint fallbackIndex = 0;
            uint maxIndex = 0;
            bool hasPoint = false;
            foreach (StringPoint point in container.Descendants<StringPoint>()) {
                uint index = point.Index?.Value ?? fallbackIndex;
                hasPoint = true;
                if (index > maxIndex) {
                    maxIndex = index;
                }

                string? value = point.NumericValue?.Text;
                values[index] = string.IsNullOrWhiteSpace(value) ? string.Empty : value!;

                fallbackIndex++;
            }

            if (!hasPoint) {
                return Array.Empty<string>();
            }

            var result = new List<string>();
            for (uint index = 0; index <= maxIndex; index++) {
                result.Add(values.TryGetValue(index, out string? value) ? value : string.Empty);
            }

            return result;
        }

        private static IReadOnlyList<double> ExtractNativeWordChartNumberValues(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<double>();
            }

            var values = new SortedDictionary<uint, double>();
            uint fallbackIndex = 0;
            uint maxIndex = 0;
            bool hasPoint = false;
            foreach (NumericPoint point in container.Descendants<NumericPoint>()) {
                uint index = point.Index?.Value ?? fallbackIndex;
                hasPoint = true;
                if (index > maxIndex) {
                    maxIndex = index;
                }

                if (TryParseNativeWordChartNumber(point.NumericValue?.Text, out double value)) {
                    values[index] = value;
                }

                fallbackIndex++;
            }

            if (!hasPoint) {
                return Array.Empty<double>();
            }

            var result = new List<double>();
            for (uint index = 0; index <= maxIndex; index++) {
                result.Add(values.TryGetValue(index, out double value) ? value : 0D);
            }

            return result;
        }

        private static IReadOnlyList<double> NormalizeNativeWordChartNumberValues(IReadOnlyList<double> values, int count, bool useIndexDefaults) {
            var result = new List<double>(count);
            for (int index = 0; index < count; index++) {
                if (index < values.Count) {
                    result.Add(values[index]);
                } else {
                    result.Add(useIndexDefaults ? index + 1D : 0D);
                }
            }

            return result;
        }

        private static bool TryParseNativeWordChartNumber(string? text, out double value) {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                if (!double.IsNaN(value) && !double.IsInfinity(value)) {
                    return true;
                }
            }

            value = 0D;
            return false;
        }

        private static string GetNativeWordChartSeriesName(OpenXmlElement seriesElement, int index) {
            string? name = GetFirstNativeWordChartText(seriesElement.Elements<SeriesText>().FirstOrDefault());
            return string.IsNullOrWhiteSpace(name)
                ? "Series " + (index + 1).ToString(CultureInfo.InvariantCulture)
                : name!;
        }

        private static string? GetNativeWordChartTitle(Chart chart) =>
            GetFirstNativeWordChartText(chart.Title);

        private static string? GetFirstNativeWordChartText(OpenXmlElement? element) {
            if (element == null) {
                return null;
            }

            string drawingText = string.Concat(element.Descendants<A.Text>().Select(text => text.Text));
            if (!string.IsNullOrWhiteSpace(drawingText)) {
                return drawingText;
            }

            string cachedText = string.Concat(element.Descendants<NumericValue>().Select(text => text.Text));
            return string.IsNullOrWhiteSpace(cachedText) ? null : cachedText;
        }

        private static string GetNativeWordChartName(WordChart chart, ChartPart? chartPart, string? title) {
            string? drawingName = chart.Drawing?.Inline?.DocProperties?.Name?.Value;
            if (!string.IsNullOrWhiteSpace(drawingName)) {
                return drawingName!;
            }

            if (!string.IsNullOrWhiteSpace(title)) {
                return title!;
            }

            return chartPart?.Uri?.ToString() ?? "Word chart";
        }

        private static string GetNativeWordChartDisplayName(OfficeChartSnapshot snapshot) =>
            string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title!;

        private static (double Width, double Height) GetNativeWordChartSizePoints(WordChart chart) {
            long? cx = chart.Drawing?.Inline?.Extent?.Cx?.Value;
            long? cy = chart.Drawing?.Inline?.Extent?.Cy?.Value;
            double width = cx.HasValue && cx.Value > 0 ? cx.Value / NativeEmusPerPoint : 360D;
            double height = cy.HasValue && cy.Value > 0 ? cy.Value / NativeEmusPerPoint : 216D;
            return (width, height);
        }

        private static OfficeChartLayout? CreateNativeWordChartLayout(OpenXmlElement chartElement) {
            DataLabels? labels = GetNativeWordChartDataLabels(chartElement);
            if (labels == null) {
                return null;
            }

            bool showValue = IsNativeWordChartBooleanOn(labels.GetFirstChild<ShowValue>());
            bool showPercent = IsNativeWordChartBooleanOn(labels.GetFirstChild<ShowPercent>());
            bool showCategoryName = IsNativeWordChartBooleanOn(labels.GetFirstChild<ShowCategoryName>());
            bool showSeriesName = IsNativeWordChartBooleanOn(labels.GetFirstChild<ShowSeriesName>());
            bool showDataLabels = showValue || showPercent || showCategoryName || showSeriesName;
            if (!showDataLabels) {
                return null;
            }

            string? separator = labels.GetFirstChild<Separator>()?.InnerText;
            return new OfficeChartLayout(
                showDataLabels: true,
                showDataLabelValues: showValue,
                showDataLabelPercentages: showPercent,
                showDataLabelCategoryNames: showCategoryName,
                showDataLabelSeriesNames: showSeriesName,
                dataLabelSeparator: separator);
        }

        private static DataLabels? GetNativeWordChartDataLabels(OpenXmlElement chartElement) {
            DataLabels? chartLabels = chartElement.GetFirstChild<DataLabels>();
            if (chartLabels != null) {
                return chartLabels;
            }

            foreach (OpenXmlElement seriesElement in chartElement.ChildElements.Where(element => element.LocalName == "ser")) {
                DataLabels? seriesLabels = seriesElement.GetFirstChild<DataLabels>();
                if (seriesLabels != null) {
                    return seriesLabels;
                }
            }

            return null;
        }

        private static bool IsNativeWordChartBooleanOn(BooleanType? value) =>
            value != null && (value.Val == null || value.Val.Value);
    }
}
