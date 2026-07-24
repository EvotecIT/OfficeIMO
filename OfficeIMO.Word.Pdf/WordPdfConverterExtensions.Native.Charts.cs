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
        private const int MaxNativeWordChartSeries = 256;
        private const int MaxNativeWordChartPoints = 4096;
        private const double MinNativeWordChartWidthPoints = 240D;
        private const double MaxNativeWordChartWidthPoints = 420D;
        private const double MinNativeWordChartHeightPoints = 150D;
        private const double MaxNativeWordChartHeightPoints = 260D;
        private const double NativeWordChartTitleTopPadding = 31D;
        private const double NativeWordChartSpacingAfter = NativeDefaultParagraphSpacingAfter;

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

            if (!string.IsNullOrWhiteSpace(warning) && options != null) {
                AddNativeExportWarning(
                    options,
                    "NativeBodyChartSimplified",
                    source,
                    warning!);
            }

            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(snapshot!);
            if (rendering.QualityReport.HasIssues && options != null) {
                AddNativeExportWarning(
                    options,
                    "NativeBodyChartQuality",
                    source,
                    "Exported Word chart '" + GetNativeWordChartDisplayName(snapshot!) + "' with shared drawing quality warnings: " + string.Join("; ", rendering.QualityReport.Issues.Select(issue => issue.ToString())));
            }

            pdf.Drawing(rendering.Drawing, align, spacingBefore: 2D, spacingAfter: NativeWordChartSpacingAfter);
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

            List<OpenXmlElement> allChartElements = plotArea.ChildElements
                .Where(IsNativeWordChartElement)
                .ToList();
            if (allChartElements.Count > 1) {
                warning = "Word combo charts are not partially exported because omitting a plot can misrepresent the source data.";
                return false;
            }

            List<OpenXmlElement> chartElements = allChartElements
                .Where(IsNativeSupportedWordChartElement)
                .ToList();
            if (chartElements.Count == 0) {
                warning = "Word chart type is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            OpenXmlElement chartElement = chartElements[0];
            if (!TryMapNativeWordChartKind(chartElement, out OfficeChartKind chartKind)) {
                warning = "Word chart type '" + chartElement.LocalName + "' is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            try {
                IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors = GetNativeDrawingThemeColors(chartPart);
                IReadOnlyList<OfficeChartSeries> series = ExtractNativeWordChartSeries(openXmlChart!, chartElement, chartKind, themeColors, out IReadOnlyList<string> categories);
                if (categories.Count == 0 || series.Count == 0) {
                    warning = "Word chart does not contain cached categories and values that can be rendered without Office.";
                    return false;
                }

                (double width, double height) = GetNativeWordChartSizePoints(chart);
                string? title = GetNativeWordChartTitle(openXmlChart!);
                string name = GetNativeWordChartName(chart, chartPart, title);
                OfficeChartStyle? style = CreateNativeWordChartStyle(openXmlChart!, chartElement, plotArea, chartKind, categories.Count, series.Count, themeColors);
                OfficeChartLayout? layout = CreateNativeWordChartLayout(openXmlChart!, chartElement, plotArea, chartKind, categories.Count);
                snapshot = new OfficeChartSnapshot(
                    name,
                    title,
                    chartKind,
                    new OfficeChartData(categories, series),
                    width,
                    height,
                    style: style,
                    layout: layout);
            } catch (NativeWordChartLimitException ex) {
                warning = ex.Message;
                return false;
            }

            return true;
        }

        private static bool IsNativeSupportedWordChartElement(OpenXmlElement element) =>
            element.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or "areaChart" or "area3DChart" or "pieChart" or "pie3DChart" or "doughnutChart" or "scatterChart" or "radarChart";

        private static bool IsNativeWordChartElement(OpenXmlElement element) =>
            element.LocalName.EndsWith("Chart", StringComparison.Ordinal);

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

        private static IReadOnlyList<OfficeChartSeries> ExtractNativeWordChartSeries(Chart chart, OpenXmlElement chartElement, OfficeChartKind chartKind, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors, out IReadOnlyList<string> categories) {
            var series = new List<OfficeChartSeries>();
            var categoryList = new List<string>();
            bool isScatter = chartKind == OfficeChartKind.Scatter;
            bool varyColorsByPoint = !isScatter && !IsNativeWordPieLikeChart(chartKind) && IsNativeWordChartVaryColorsEnabled(chartElement);
            HashSet<uint> hiddenLegendIndexes = GetNativeWordHiddenLegendIndexes(chart);

            int seriesIndex = 0;
            foreach ((OpenXmlElement seriesElement, int originalSeriesIndex) in GetNativeWordOrderedSeriesElements(chartElement)) {
                if (seriesIndex >= MaxNativeWordChartSeries) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
                }

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
                } else if (currentCategories.Count > categoryList.Count) {
                    for (int index = categoryList.Count; index < currentCategories.Count; index++) {
                        categoryList.Add(currentCategories[index]);
                    }
                } else if (values.Count > categoryList.Count) {
                    for (int index = categoryList.Count; index < values.Count; index++) {
                        categoryList.Add("Category " + (index + 1).ToString(CultureInfo.InvariantCulture));
                    }
                }

                IReadOnlyList<OfficeColor?>? pointColors = ExtractNativeWordChartPointColors(seriesElement, values.Count, themeColors);
                if (pointColors == null && varyColorsByPoint && seriesIndex == 0) {
                    pointColors = CreateNativeWordChartVaryPointColors(values.Count);
                }

                OfficeColor? seriesColor = null;
                if (IsNativeWordPieLikeChart(chartKind) &&
                    TryGetNativeWordChartFillColor(seriesElement, themeColors, out OfficeColor pieLikeSeriesColor)) {
                    seriesColor = pieLikeSeriesColor;
                }

                series.Add(new OfficeChartSeries(
                    GetNativeWordChartSeriesName(seriesElement, seriesIndex),
                    values,
                    xValues,
                    seriesColor,
                    pointColors,
                    !IsNativeWordChartSeriesMarkerHidden(seriesElement),
                    !hiddenLegendIndexes.Contains((uint)originalSeriesIndex),
                    !IsNativeWordLineLikeChart(chartKind) || !HasNativeDrawingOutlineNoFill(seriesElement.GetFirstChild<ChartShapeProperties>())));
                seriesIndex++;
            }

            categories = categoryList;
            return series;
        }

        private static IReadOnlyList<string> ExtractNativeWordChartCategories(OpenXmlElement? categoryAxisData, int valueCount) {
            EnsureNativeWordChartPointCount(valueCount);
            List<string> categories = ExtractNativeWordChartStringValues(categoryAxisData).ToList();
            if (categories.Count == 0) {
                categories.AddRange(ExtractNativeWordChartNumberValues(categoryAxisData).Select(value => value.ToString("0.####", CultureInfo.InvariantCulture)));
            }

            if (categories.Count == 0 && valueCount > 0) {
                categories.AddRange(CreateNativeWordChartDefaultCategories(valueCount));
            }

            for (int index = categories.Count; index < valueCount; index++) {
                categories.Add("Category " + (index + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static IReadOnlyList<string> CreateNativeWordChartDefaultCategories(int count) {
            EnsureNativeWordChartPointCount(count);
            var categories = new List<string>();
            for (int i = 0; i < count; i++) {
                categories.Add("Category " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static bool IsNativeWordChartVaryColorsEnabled(OpenXmlElement chartElement) =>
            IsNativeWordChartBooleanOn(chartElement.GetFirstChild<VaryColors>());

        private static bool IsNativeWordChartVaryColorsDisabled(OpenXmlElement chartElement) {
            VaryColors? varyColors = chartElement.GetFirstChild<VaryColors>();
            return varyColors?.Val != null && !varyColors.Val.Value;
        }

        private static HashSet<uint> GetNativeWordHiddenLegendIndexes(Chart chart) {
            var indexes = new HashSet<uint>();
            Legend? legend = chart.GetFirstChild<Legend>();
            if (legend == null) {
                return indexes;
            }

            foreach (LegendEntry entry in legend.Elements<LegendEntry>()) {
                uint? index = entry.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>()?.Val?.Value;
                if (index.HasValue && IsNativeWordChartBooleanOn(entry.GetFirstChild<Delete>())) {
                    indexes.Add(index.Value);
                }
            }

            return indexes;
        }

        private static IReadOnlyList<OfficeColor?> CreateNativeWordChartVaryPointColors(int valueCount) {
            var colors = new OfficeColor?[valueCount];
            for (int index = 0; index < colors.Length; index++) {
                colors[index] = OfficeChartStyle.Default.GetSeriesColor(index);
            }

            return colors;
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
                if (fallbackIndex >= MaxNativeWordChartPoints) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported point count.");
                }

                uint index = point.Index?.Value ?? fallbackIndex;
                EnsureNativeWordChartPointIndex(index);
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
                result.Add(values.TryGetValue(index, out string? value)
                    ? value
                    : "Category " + (index + 1U).ToString(CultureInfo.InvariantCulture));
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
                if (fallbackIndex >= MaxNativeWordChartPoints) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported point count.");
                }

                uint index = point.Index?.Value ?? fallbackIndex;
                EnsureNativeWordChartPointIndex(index);
                hasPoint = true;
                if (index > maxIndex) {
                    maxIndex = index;
                }

                if (TryParseNativeWordChartNumber(point.NumericValue?.Text, out double value)) {
                    values[index] = value;
                } else {
                    values[index] = double.NaN;
                }

                fallbackIndex++;
            }

            if (!hasPoint) {
                return Array.Empty<double>();
            }

            var result = new List<double>();
            for (uint index = 0; index <= maxIndex; index++) {
                result.Add(values.TryGetValue(index, out double value) ? value : double.NaN);
            }

            return result;
        }

        private static IReadOnlyList<double> NormalizeNativeWordChartNumberValues(IReadOnlyList<double> values, int count, bool useIndexDefaults) {
            EnsureNativeWordChartPointCount(count);
            var result = new List<double>(count);
            for (int index = 0; index < count; index++) {
                if (index < values.Count) {
                    result.Add(values[index]);
                } else {
                    result.Add(useIndexDefaults ? index + 1D : double.NaN);
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

        private static IReadOnlyList<(OpenXmlElement SeriesElement, int OriginalSeriesIndex)> GetNativeWordOrderedSeriesElements(OpenXmlElement chartElement) {
            var series = new List<(OpenXmlElement SeriesElement, int OriginalSeriesIndex, int Order)>();
            int index = 0;
            foreach (OpenXmlElement element in chartElement.ChildElements.Where(element => element.LocalName == "ser")) {
                if (series.Count >= MaxNativeWordChartSeries) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
                }

                series.Add((element, index, GetNativeWordChartSeriesOrder(element, index)));
                index++;
            }

            return series
                .OrderBy(item => item.Order)
                .ThenBy(item => item.OriginalSeriesIndex)
                .Select(item => (item.SeriesElement, item.OriginalSeriesIndex))
                .ToArray();
        }

        private static int GetNativeWordChartSeriesOrder(OpenXmlElement seriesElement, int fallback) {
            OpenXmlElement? order = seriesElement.ChildElements.FirstOrDefault(element => element.LocalName == "order");
            string? value = order == null ? null : GetNativeOpenXmlAttribute(order, "val");
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : fallback;
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
            long? cx = chart.Drawing?.Inline?.Extent?.Cx?.Value ?? chart.Drawing?.Anchor?.Extent?.Cx?.Value;
            long? cy = chart.Drawing?.Inline?.Extent?.Cy?.Value ?? chart.Drawing?.Anchor?.Extent?.Cy?.Value;
            double width = cx.HasValue && cx.Value > 0 ? cx.Value / NativeEmusPerPoint : 360D;
            double height = cy.HasValue && cy.Value > 0 ? cy.Value / NativeEmusPerPoint : 216D;
            width = Math.Min(MaxNativeWordChartWidthPoints, Math.Max(MinNativeWordChartWidthPoints, width));
            height = Math.Min(MaxNativeWordChartHeightPoints, Math.Max(MinNativeWordChartHeightPoints, height));
            return (width, height);
        }

        private static OfficeChartStyle? CreateNativeWordChartStyle(Chart chart, OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind, int categoryCount, int seriesCount, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            EnsureNativeWordChartPointCount(categoryCount);
            if (seriesCount > MaxNativeWordChartSeries) {
                throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
            }

            int paletteCount = IsNativeWordPieLikeChart(chartKind) ? categoryCount : seriesCount;
            if (paletteCount <= 0) {
                return null;
            }

            var palette = new List<OfficeColor>(paletteCount);
            for (int index = 0; index < paletteCount; index++) {
                palette.Add(OfficeChartStyle.Default.GetSeriesColor(index));
            }

            if (IsNativeWordPieLikeChart(chartKind)) {
                OpenXmlElement? seriesElement = chartElement.ChildElements.FirstOrDefault(element => element.LocalName == "ser" && IsNativeWordChartSeriesRenderable(element, chartKind));
                if (TryGetNativeWordChartFillColor(seriesElement, themeColors, out OfficeColor seriesColor)) {
                    for (int index = 0; index < palette.Count; index++) {
                        palette[index] = seriesColor;
                    }

                } else if (IsNativeWordChartVaryColorsDisabled(chartElement) && palette.Count > 0) {
                    OfficeColor singleColor = palette[0];
                    for (int index = 0; index < palette.Count; index++) {
                        palette[index] = singleColor;
                    }

                }

                foreach (DataPoint point in seriesElement?.Elements<DataPoint>() ?? Enumerable.Empty<DataPoint>()) {
                    uint? pointIndex = point.Index?.Val?.Value;
                    if (!pointIndex.HasValue || pointIndex.Value >= (uint)palette.Count) {
                        continue;
                    }

                    int index = (int)pointIndex.Value;
                    if (TryGetNativeWordChartFillColor(point, themeColors, out OfficeColor pointColor)) {
                        palette[index] = pointColor;
                    }
                }
            } else {
                int index = 0;
                foreach ((OpenXmlElement seriesElement, _) in GetNativeWordOrderedSeriesElements(chartElement).Where(item => IsNativeWordChartSeriesRenderable(item.SeriesElement, chartKind))) {
                    if (index >= palette.Count) {
                        break;
                    }

                    if (TryGetNativeWordChartSeriesColor(seriesElement, chartKind, themeColors, out OfficeColor seriesColor)) {
                        palette[index] = seriesColor;
                    }

                    index++;
                }
            }

            OfficeColor? backgroundColor = null;
            OfficeColor? borderColor = null;
            ChartShapeProperties? chartShape = chart.GetFirstChild<ChartShapeProperties>();
            bool showBackground = !HasNativeDrawingNoFill(chartShape);
            bool hasExplicitChartNoFill = chartShape != null && !showBackground;
            bool showBorder = !HasNativeDrawingOutlineNoFill(chartShape);
            bool hasExplicitChartNoLine = chartShape != null && !showBorder;
            if (TryGetNativeDrawingSolidFillColor(chartShape, out OfficeColor chartFill, themeColors)) {
                backgroundColor = chartFill;
            }

            if (TryGetNativeDrawingOutlineColor(chartShape, out OfficeColor chartBorder, themeColors)) {
                borderColor = chartBorder;
            }

            OfficeColor? plotAreaBackgroundColor = null;
            OfficeColor? plotAreaBorderColor = null;
            ChartShapeProperties? plotShape = plotArea.GetFirstChild<ChartShapeProperties>();
            if (TryGetNativeDrawingSolidFillColor(plotShape, out OfficeColor plotFill, themeColors)) {
                plotAreaBackgroundColor = plotFill;
            }

            if (TryGetNativeDrawingOutlineColor(plotShape, out OfficeColor plotBorder, themeColors)) {
                plotAreaBorderColor = plotBorder;
            }

            OfficeColor? axisColor = GetNativeWordChartAxisLineColor(chartElement, plotArea, themeColors);
            OfficeColor? gridLineColor = GetNativeWordChartMajorGridLineColor(chartElement, plotArea, themeColors);
            bool showGridLines = HasNativeWordChartMajorGridLines(chartElement, plotArea);
            OfficeColor? titleColor = GetNativeWordChartTitleColor(chart, themeColors);

            return new OfficeChartStyle(
                showBackground: showBackground,
                palette: palette,
                fontFamily: "Calibri",
                backgroundColor: backgroundColor ?? OfficeColor.White,
                borderColor: borderColor ?? OfficeColor.FromRgb(127, 127, 127),
                axisColor: axisColor ?? OfficeColor.Black,
                gridLineColor: gridLineColor ?? OfficeColor.FromRgb(217, 217, 217),
                textColor: OfficeColor.Black,
                mutedTextColor: OfficeColor.Black,
                titleColor: titleColor ?? OfficeColor.Black,
                plotAreaBackgroundColor: plotAreaBackgroundColor,
                plotAreaBorderColor: plotAreaBorderColor,
                showGridLines: showGridLines,
                showBorder: showBorder);
        }

        private static bool IsNativeWordChartSeriesRenderable(OpenXmlElement seriesElement, OfficeChartKind chartKind) {
            OpenXmlElement? valuesElement = chartKind == OfficeChartKind.Scatter
                ? seriesElement.Elements<YValues>().FirstOrDefault()
                : seriesElement.Elements<Values>().FirstOrDefault();
            return ExtractNativeWordChartNumberValues(valuesElement).Count > 0;
        }

        private static OfficeColor? GetNativeWordChartTitleColor(Chart chart, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            Title? title = chart.Title;
            if (title == null) {
                return null;
            }

            foreach (OpenXmlElement textProperties in title.Descendants().Where(element => element.LocalName == "defRPr" || element.LocalName == "rPr")) {
                if (TryGetNativeDrawingSolidFillColor(textProperties, out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            return null;
        }

        private static OfficeColor? GetNativeWordChartAxisLineColor(OpenXmlElement chartElement, PlotArea plotArea, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            if (TryGetNativeWordChartAxisLineColor<ValueAxis>(chartElement, plotArea, themeColors, out OfficeColor valueAxisColor)) {
                return valueAxisColor;
            }

            if (TryGetNativeWordChartAxisLineColor<CategoryAxis>(chartElement, plotArea, themeColors, out OfficeColor categoryAxisColor)) {
                return categoryAxisColor;
            }

            return null;
        }

        private static bool TryGetNativeWordChartAxisLineColor<TAxis>(OpenXmlElement chartElement, PlotArea plotArea, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors, out OfficeColor color)
            where TAxis : OpenXmlElement {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (TAxis axis in plotArea.Elements<TAxis>()) {
                uint? axisId = axis.GetFirstChild<AxisId>()?.Val?.Value;
                if (axisId.HasValue &&
                    chartAxisIds.Contains(axisId.Value) &&
                    TryGetNativeDrawingOutlineColor(axis.GetFirstChild<ChartShapeProperties>(), out color, themeColors)) {
                    return true;
                }
            }

            foreach (TAxis axis in plotArea.Elements<TAxis>()) {
                if (TryGetNativeDrawingOutlineColor(axis.GetFirstChild<ChartShapeProperties>(), out color, themeColors)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool HasNativeWordChartAxisLineNoFill<TAxis>(OpenXmlElement chartElement, PlotArea plotArea)
            where TAxis : OpenXmlElement {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (TAxis axis in plotArea.Elements<TAxis>()) {
                uint? axisId = axis.GetFirstChild<AxisId>()?.Val?.Value;
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    return HasNativeDrawingOutlineNoFill(axis.GetFirstChild<ChartShapeProperties>());
                }
            }

            TAxis? fallback = plotArea.Elements<TAxis>().FirstOrDefault();
            return HasNativeDrawingOutlineNoFill(fallback?.GetFirstChild<ChartShapeProperties>());
        }

        private static OfficeColor? GetNativeWordChartMajorGridLineColor(OpenXmlElement chartElement, PlotArea plotArea, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                uint? axisId = axis.AxisId?.Val?.Value;
                if (axisId.HasValue &&
                    chartAxisIds.Contains(axisId.Value) &&
                    TryGetNativeDrawingOutlineColor(axis.GetFirstChild<MajorGridlines>()?.GetFirstChild<ChartShapeProperties>(), out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                if (TryGetNativeDrawingOutlineColor(axis.GetFirstChild<MajorGridlines>()?.GetFirstChild<ChartShapeProperties>(), out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            return null;
        }

        private static bool HasNativeWordChartMajorGridLines(OpenXmlElement chartElement, PlotArea plotArea) {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            bool hasMatchingValueAxis = false;
            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                uint? axisId = axis.AxisId?.Val?.Value;
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    hasMatchingValueAxis = true;
                    MajorGridlines? gridlines = axis.GetFirstChild<MajorGridlines>();
                    if (gridlines != null) {
                        return !HasNativeDrawingOutlineNoFill(gridlines.GetFirstChild<ChartShapeProperties>());
                    }
                }
            }

            if (hasMatchingValueAxis) {
                return false;
            }

            bool hasAnyValueAxis = false;
            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                hasAnyValueAxis = true;
                MajorGridlines? gridlines = axis.GetFirstChild<MajorGridlines>();
                if (gridlines != null) {
                    return !HasNativeDrawingOutlineNoFill(gridlines.GetFirstChild<ChartShapeProperties>());
                }
            }

            return !hasAnyValueAxis;
        }

        private static bool IsNativeWordPieLikeChart(OfficeChartKind chartKind) =>
            chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut;

        private static bool TryGetNativeWordChartFillColor(OpenXmlElement? element, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors, out OfficeColor color) {
            return TryGetNativeDrawingSolidFillColor(element?.GetFirstChild<ChartShapeProperties>(), out color, themeColors);
        }

        private static bool TryGetNativeWordChartSeriesColor(OpenXmlElement? element, OfficeChartKind chartKind, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors, out OfficeColor color) {
            ChartShapeProperties? properties = element?.GetFirstChild<ChartShapeProperties>();
            if (TryGetNativeDrawingSolidFillColor(properties, out color, themeColors)) {
                return true;
            }

            return IsNativeWordLineLikeChart(chartKind) && TryGetNativeDrawingOutlineColor(properties, out color, themeColors);
        }

        private static bool IsNativeWordLineLikeChart(OfficeChartKind chartKind) =>
            chartKind == OfficeChartKind.Line ||
            chartKind == OfficeChartKind.LineStacked ||
            chartKind == OfficeChartKind.LineStacked100 ||
            chartKind == OfficeChartKind.Scatter ||
            chartKind == OfficeChartKind.Radar;

        private static IReadOnlyList<OfficeColor?>? ExtractNativeWordChartPointColors(OpenXmlElement seriesElement, int valueCount, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            if (valueCount <= 0) {
                return null;
            }

            OfficeColor?[] colors = new OfficeColor?[valueCount];
            bool anyColor = false;
            foreach (DataPoint point in seriesElement.Elements<DataPoint>()) {
                uint? index = point.Index?.Val?.Value;
                if (!index.HasValue || index.Value >= valueCount) {
                    continue;
                }

                if (TryGetNativeWordChartFillColor(point, themeColors, out OfficeColor color)) {
                    colors[index.Value] = color;
                    anyColor = true;
                }
            }

            return anyColor ? colors : null;
        }

        private static OfficeChartLayout? CreateNativeWordChartLayout(Chart chart, OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind, int categoryCount) {
            NativeWordChartDataLabelState dataLabels = GetNativeWordChartDataLabels(chartElement, chartKind);
            DataLabels? labels = dataLabels.Labels;
            bool showValue = dataLabels.ShowValue || HasNativeWordChartDataLabelFlag<ShowValue>(labels);
            bool showPercent = labels != null && IsNativeWordPieLikeChart(chartKind) && (dataLabels.ShowPercent || HasNativeWordChartDataLabelFlag<ShowPercent>(labels));
            bool showCategoryName = dataLabels.ShowCategoryName || HasNativeWordChartDataLabelFlag<ShowCategoryName>(labels);
            bool showSeriesName = dataLabels.ShowSeriesName || HasNativeWordChartDataLabelFlag<ShowSeriesName>(labels);
            bool showDataLabels = showValue || showPercent || showCategoryName || showSeriesName;
            bool showLegend = HasNativeWordChartLegend(chart);
            OfficeChartLegendPosition legendPosition = GetNativeWordChartLegendPosition(chart);
            bool overlayLegend = IsNativeWordChartLegendOverlay(chart);
            OfficeChartDataLabelPosition dataLabelPosition = GetNativeWordChartDataLabelPosition(labels);
            string? dataLabelNumberFormat = GetNativeWordChartDataLabelNumberFormat(labels);
            bool showMarkers = !AreNativeWordChartMarkersHidden(chartElement, chartKind);
            string? axisNumberFormat = GetNativeWordChartValueAxisNumberFormat(chartElement, plotArea);
            string? categoryAxisTitle = GetNativeWordChartCategoryAxisTitle(chartElement, plotArea);
            string? valueAxisTitle = GetNativeWordChartValueAxisTitle(chartElement, plotArea);
            string? horizontalAxisNumberFormat = axisNumberFormat;
            string? verticalAxisNumberFormat = axisNumberFormat;
            bool connectScatterPoints = !IsNativeWordMarkerOnlyScatter(chartElement, chartKind);
            bool fillRadarSeries = chartKind != OfficeChartKind.Radar || IsNativeWordFilledRadar(chartElement, chartKind);
            bool showCategoryAxis = IsNativeWordChartCategoryAxisVisible(chartElement, plotArea, chartKind);
            bool showValueAxis = IsNativeWordChartValueAxisVisible(chartElement, plotArea, chartKind);
            bool showCategoryAxisLine = showCategoryAxis && !HasNativeWordChartAxisLineNoFill<CategoryAxis>(chartElement, plotArea);
            bool showValueAxisLine = showValueAxis && !HasNativeWordChartAxisLineNoFill<ValueAxis>(chartElement, plotArea);
            bool showCategoryAxisLabels = showCategoryAxis && IsNativeWordChartCategoryAxisLabelsVisible(chartElement, plotArea, chartKind);
            bool showValueAxisLabels = showValueAxis && IsNativeWordChartValueAxisLabelsVisible(chartElement, plotArea, chartKind);
            bool overlayTitle = IsNativeWordChartTitleOverlay(chart);
            double? titleTopPadding = !overlayTitle && !string.IsNullOrWhiteSpace(GetNativeWordChartTitle(chart))
                ? NativeWordChartTitleTopPadding
                : null;
            if (chartKind == OfficeChartKind.Scatter) {
                NativeWordScatterAxisMetadata scatterAxisMetadata = GetNativeWordScatterAxisMetadata(chartElement, plotArea);
                horizontalAxisNumberFormat = scatterAxisMetadata.HorizontalNumberFormat ?? axisNumberFormat;
                verticalAxisNumberFormat = scatterAxisMetadata.VerticalNumberFormat ?? axisNumberFormat;
                axisNumberFormat = verticalAxisNumberFormat ?? horizontalAxisNumberFormat;
                categoryAxisTitle = scatterAxisMetadata.HorizontalTitle ?? categoryAxisTitle;
                valueAxisTitle = scatterAxisMetadata.VerticalTitle ?? valueAxisTitle;
                showCategoryAxis = scatterAxisMetadata.HorizontalVisible;
                showValueAxis = scatterAxisMetadata.VerticalVisible;
                showCategoryAxisLine = scatterAxisMetadata.HorizontalLineVisible;
                showValueAxisLine = scatterAxisMetadata.VerticalLineVisible;
                showCategoryAxisLabels = scatterAxisMetadata.HorizontalLabelsVisible;
                showValueAxisLabels = scatterAxisMetadata.VerticalLabelsVisible;
            }
            int? maximumCategoryAxisLabels = null;
            int? maximumHorizontalCategoryAxisLabels = null;
            int? maximumRadarCategoryLabels = null;
            int? maximumLabelsFromAxisSkip = GetNativeWordChartMaximumCategoryAxisLabels(chartElement, plotArea, categoryCount);
            if (maximumLabelsFromAxisSkip.HasValue) {
                if (IsNativeWordHorizontalBarChart(chartKind)) {
                    maximumHorizontalCategoryAxisLabels = maximumLabelsFromAxisSkip;
                } else if (chartKind == OfficeChartKind.Radar) {
                    maximumRadarCategoryLabels = maximumLabelsFromAxisSkip;
                } else {
                    maximumCategoryAxisLabels = maximumLabelsFromAxisSkip;
                }
            }

            if (showLegend &&
                legendPosition == OfficeChartLegendPosition.Right &&
                !overlayLegend &&
                !showDataLabels &&
                !maximumCategoryAxisLabels.HasValue &&
                !maximumHorizontalCategoryAxisLabels.HasValue &&
                !maximumRadarCategoryLabels.HasValue &&
                string.IsNullOrWhiteSpace(axisNumberFormat) &&
                string.IsNullOrWhiteSpace(horizontalAxisNumberFormat) &&
                string.IsNullOrWhiteSpace(verticalAxisNumberFormat) &&
                string.IsNullOrWhiteSpace(categoryAxisTitle) &&
                string.IsNullOrWhiteSpace(valueAxisTitle) &&
                connectScatterPoints &&
                fillRadarSeries &&
                showCategoryAxis &&
                showValueAxis &&
                showCategoryAxisLine &&
                showValueAxisLine &&
                showCategoryAxisLabels &&
                showValueAxisLabels &&
                !overlayTitle &&
                !titleTopPadding.HasValue) {
                return null;
            }

            string? separator = labels?.GetFirstChild<Separator>()?.InnerText;
            HashSet<uint> hiddenLegendIndexes = GetNativeWordHiddenLegendIndexes(chart);
            int[]? hiddenCategoryLegendIndexes = IsNativeWordPieLikeChart(chartKind) && hiddenLegendIndexes.Count > 0
                ? hiddenLegendIndexes
                    .Where(index => index <= int.MaxValue)
                    .Select(index => (int)index)
                    .ToArray()
                : null;
            var layout = new OfficeChartLayout(
                maximumCategoryAxisLabels: maximumCategoryAxisLabels,
                maximumHorizontalCategoryAxisLabels: maximumHorizontalCategoryAxisLabels,
                maximumRadarCategoryLabels: maximumRadarCategoryLabels,
                showLegend: showLegend,
                legendPosition: legendPosition,
                overlayLegend: overlayLegend,
                showDataLabels: showDataLabels,
                showDataLabelValues: showValue,
                showDataLabelPercentages: showPercent,
                showDataLabelCategoryNames: showCategoryName,
                showDataLabelSeriesNames: showSeriesName,
                dataLabelSeparator: separator,
                dataLabelPosition: dataLabelPosition,
                dataLabelNumberFormat: dataLabelNumberFormat,
                showMarkers: showMarkers,
                axisNumberFormat: axisNumberFormat,
                categoryAxisTitle: categoryAxisTitle,
                valueAxisTitle: valueAxisTitle,
                horizontalAxisNumberFormat: horizontalAxisNumberFormat,
                verticalAxisNumberFormat: verticalAxisNumberFormat,
                connectScatterPoints: connectScatterPoints,
                fillRadarSeries: fillRadarSeries,
                showCategoryAxis: showCategoryAxis,
                showValueAxis: showValueAxis,
                showCategoryAxisLine: showCategoryAxisLine,
                showValueAxisLine: showValueAxisLine,
                showCategoryAxisLabels: showCategoryAxisLabels,
                showValueAxisLabels: showValueAxisLabels,
                overlayTitle: overlayTitle,
                titleTopPadding: titleTopPadding,
                dataLabelSeriesIndexes: dataLabels.SeriesIndexes,
                dataLabelPointIndexes: dataLabels.PointIndexes,
                hiddenDataLabelPointIndexes: dataLabels.HiddenPointIndexes,
                hiddenCategoryLegendIndexes: hiddenCategoryLegendIndexes);

            return layout;
        }

        private static bool IsNativeWordMarkerOnlyScatter(OpenXmlElement chartElement, OfficeChartKind chartKind) =>
            chartKind == OfficeChartKind.Scatter &&
            chartElement.GetFirstChild<ScatterStyle>()?.Val?.Value == ScatterStyleValues.Marker;

        private static bool IsNativeWordFilledRadar(OpenXmlElement chartElement, OfficeChartKind chartKind) =>
            chartKind == OfficeChartKind.Radar &&
            chartElement.GetFirstChild<RadarStyle>()?.Val?.Value == RadarStyleValues.Filled;

        private readonly struct NativeWordScatterAxisMetadata {
            public NativeWordScatterAxisMetadata(
                string? horizontalNumberFormat,
                string? verticalNumberFormat,
                string? horizontalTitle,
                string? verticalTitle,
                bool horizontalVisible,
                bool verticalVisible,
                bool horizontalLineVisible,
                bool verticalLineVisible,
                bool horizontalLabelsVisible,
                bool verticalLabelsVisible) {
                HorizontalNumberFormat = horizontalNumberFormat;
                VerticalNumberFormat = verticalNumberFormat;
                HorizontalTitle = horizontalTitle;
                VerticalTitle = verticalTitle;
                HorizontalVisible = horizontalVisible;
                VerticalVisible = verticalVisible;
                HorizontalLineVisible = horizontalLineVisible;
                VerticalLineVisible = verticalLineVisible;
                HorizontalLabelsVisible = horizontalLabelsVisible;
                VerticalLabelsVisible = verticalLabelsVisible;
            }

            public string? HorizontalNumberFormat { get; }
            public string? VerticalNumberFormat { get; }
            public string? HorizontalTitle { get; }
            public string? VerticalTitle { get; }
            public bool HorizontalVisible { get; }
            public bool VerticalVisible { get; }
            public bool HorizontalLineVisible { get; }
            public bool VerticalLineVisible { get; }
            public bool HorizontalLabelsVisible { get; }
            public bool VerticalLabelsVisible { get; }
        }

        private static NativeWordScatterAxisMetadata GetNativeWordScatterAxisMetadata(OpenXmlElement chartElement, PlotArea plotArea) {
            IReadOnlyList<uint> valueAxisIds = GetNativeWordChartAxisIds(chartElement)
                .Where(axisId => GetNativeWordChartValueAxis(plotArea, axisId) != null)
                .ToArray();

            ValueAxis? horizontalAxis = valueAxisIds.Count > 0 ? GetNativeWordChartValueAxis(plotArea, valueAxisIds[0]) : null;
            ValueAxis? verticalAxis = valueAxisIds.Count > 1 ? GetNativeWordChartValueAxis(plotArea, valueAxisIds[1]) : null;
            return new NativeWordScatterAxisMetadata(
                GetNativeWordChartAxisNumberFormat(horizontalAxis),
                GetNativeWordChartAxisNumberFormat(verticalAxis),
                horizontalAxis == null ? null : GetNativeWordChartAxisTitle(horizontalAxis),
                verticalAxis == null ? null : GetNativeWordChartAxisTitle(verticalAxis),
                !IsNativeWordChartAxisDeleted(horizontalAxis),
                !IsNativeWordChartAxisDeleted(verticalAxis),
                !IsNativeWordChartAxisDeleted(horizontalAxis) && !HasNativeDrawingOutlineNoFill(horizontalAxis?.GetFirstChild<ChartShapeProperties>()),
                !IsNativeWordChartAxisDeleted(verticalAxis) && !HasNativeDrawingOutlineNoFill(verticalAxis?.GetFirstChild<ChartShapeProperties>()),
                !IsNativeWordChartAxisDeleted(horizontalAxis) && !IsNativeWordChartTickLabelsHidden(horizontalAxis),
                !IsNativeWordChartAxisDeleted(verticalAxis) && !IsNativeWordChartTickLabelsHidden(verticalAxis));
        }

        private static string? GetNativeWordChartCategoryAxisTitle(OpenXmlElement chartElement, PlotArea plotArea) =>
            GetNativeWordChartAxisTitle<CategoryAxis>(chartElement, plotArea);

        private static string? GetNativeWordChartValueAxisTitle(OpenXmlElement chartElement, PlotArea plotArea) =>
            GetNativeWordChartAxisTitle<ValueAxis>(chartElement, plotArea);

        private static string? GetNativeWordChartAxisTitle<TAxis>(OpenXmlElement chartElement, PlotArea plotArea)
            where TAxis : OpenXmlElement {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (TAxis axis in plotArea.Elements<TAxis>()) {
                uint? axisId = axis.GetFirstChild<AxisId>()?.Val?.Value;
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    string? title = GetNativeWordChartAxisTitle(axis);
                    if (!string.IsNullOrWhiteSpace(title)) {
                        return title;
                    }
                }
            }

            foreach (TAxis axis in plotArea.Elements<TAxis>()) {
                string? title = GetNativeWordChartAxisTitle(axis);
                if (!string.IsNullOrWhiteSpace(title)) {
                    return title;
                }
            }

            return null;
        }

        private static string? GetNativeWordChartAxisTitle(OpenXmlElement axis) =>
            GetFirstNativeWordChartText(axis.GetFirstChild<Title>());

        private static IReadOnlyList<uint> GetNativeWordChartAxisIds(OpenXmlElement chartElement) =>
            chartElement.Elements<AxisId>()
                .Select(axis => axis.Val?.Value)
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToArray();

        private static ValueAxis? GetNativeWordChartValueAxis(PlotArea plotArea, uint axisId) =>
            plotArea.Elements<ValueAxis>().FirstOrDefault(axis => axis.AxisId?.Val?.Value == axisId);

        private static CategoryAxis? GetNativeWordChartCategoryAxis(PlotArea plotArea, uint axisId) =>
            plotArea.Elements<CategoryAxis>().FirstOrDefault(axis => axis.AxisId?.Val?.Value == axisId);

        private static bool IsNativeWordChartCategoryAxisVisible(OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind) {
            if (chartKind == OfficeChartKind.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                CategoryAxis? axis = GetNativeWordChartCategoryAxis(plotArea, axisId);
                if (axis != null) {
                    return !IsNativeWordChartAxisDeleted(axis);
                }
            }

            CategoryAxis? fallback = plotArea.Elements<CategoryAxis>().FirstOrDefault();
            return !IsNativeWordChartAxisDeleted(fallback);
        }

        private static bool IsNativeWordChartValueAxisVisible(OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind) {
            if (chartKind == OfficeChartKind.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                ValueAxis? axis = GetNativeWordChartValueAxis(plotArea, axisId);
                if (axis != null) {
                    return !IsNativeWordChartAxisDeleted(axis);
                }
            }

            ValueAxis? fallback = plotArea.Elements<ValueAxis>().FirstOrDefault();
            return !IsNativeWordChartAxisDeleted(fallback);
        }

        private static bool IsNativeWordChartAxisDeleted(OpenXmlElement? axis) =>
            axis != null && IsNativeWordChartBooleanOn(axis.GetFirstChild<Delete>());

        private static bool IsNativeWordChartCategoryAxisLabelsVisible(OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind) {
            if (chartKind == OfficeChartKind.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                CategoryAxis? axis = GetNativeWordChartCategoryAxis(plotArea, axisId);
                if (axis != null) {
                    return !IsNativeWordChartTickLabelsHidden(axis);
                }
            }

            return !IsNativeWordChartTickLabelsHidden(plotArea.Elements<CategoryAxis>().FirstOrDefault());
        }

        private static bool IsNativeWordChartValueAxisLabelsVisible(OpenXmlElement chartElement, PlotArea plotArea, OfficeChartKind chartKind) {
            if (chartKind == OfficeChartKind.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                ValueAxis? axis = GetNativeWordChartValueAxis(plotArea, axisId);
                if (axis != null) {
                    return !IsNativeWordChartTickLabelsHidden(axis);
                }
            }

            return !IsNativeWordChartTickLabelsHidden(plotArea.Elements<ValueAxis>().FirstOrDefault());
        }

        private static bool IsNativeWordChartTickLabelsHidden(OpenXmlElement? axis) {
            OpenXmlElement? position = axis?.ChildElements.FirstOrDefault(element => element.LocalName == "tickLblPos");
            return position != null && string.Equals(GetNativeOpenXmlAttribute(position, "val"), "none", StringComparison.OrdinalIgnoreCase);
        }

        private static string? GetNativeWordChartAxisNumberFormat(ValueAxis? axis) {
            string? format = axis?.GetFirstChild<NumberingFormat>()?.FormatCode?.Value;
            return string.IsNullOrWhiteSpace(format) ? null : format;
        }

        private static string? GetNativeWordChartValueAxisNumberFormat(OpenXmlElement chartElement, PlotArea plotArea) {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                uint? axisId = axis.AxisId?.Val?.Value;
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    string? format = GetNativeWordChartAxisNumberFormat(axis);
                    if (!string.IsNullOrWhiteSpace(format)) {
                        return format;
                    }
                }
            }

            foreach (ValueAxis axis in plotArea.Elements<ValueAxis>()) {
                string? format = GetNativeWordChartAxisNumberFormat(axis);
                if (!string.IsNullOrWhiteSpace(format)) {
                    return format;
                }
            }

            return null;
        }

        private static bool AreNativeWordChartMarkersHidden(OpenXmlElement chartElement, OfficeChartKind chartKind) {
            if (chartKind != OfficeChartKind.Line &&
                chartKind != OfficeChartKind.LineStacked &&
                chartKind != OfficeChartKind.LineStacked100 &&
                chartKind != OfficeChartKind.Scatter &&
                chartKind != OfficeChartKind.Radar) {
                return false;
            }

            bool sawSeries = false;
            bool sawHiddenMarker = false;
            foreach (OpenXmlElement seriesElement in chartElement.ChildElements.Where(element => element.LocalName == "ser")) {
                sawSeries = true;
                if (!IsNativeWordChartSeriesMarkerHidden(seriesElement)) {
                    return false;
                }

                sawHiddenMarker = true;
            }

            return sawSeries && sawHiddenMarker;
        }

        private static bool IsNativeWordChartSeriesMarkerHidden(OpenXmlElement seriesElement) =>
            seriesElement.GetFirstChild<Marker>()?.Symbol?.Val?.Value == MarkerStyleValues.None;

        private static string? GetNativeWordChartDataLabelNumberFormat(DataLabels? labels) {
            string? formatCode = labels?.GetFirstChild<NumberingFormat>()?.FormatCode?.Value;
            return string.IsNullOrWhiteSpace(formatCode) ? null : formatCode;
        }

        private static OfficeChartDataLabelPosition GetNativeWordChartDataLabelPosition(DataLabels? labels) {
            DataLabelPositionValues? position = labels?.GetFirstChild<DataLabelPosition>()?.Val?.Value;
            if (position == DataLabelPositionValues.Center) {
                return OfficeChartDataLabelPosition.Center;
            }

            if (position == DataLabelPositionValues.InsideBase) {
                return OfficeChartDataLabelPosition.InsideBase;
            }

            if (position == DataLabelPositionValues.InsideEnd) {
                return OfficeChartDataLabelPosition.InsideEnd;
            }

            if (position == DataLabelPositionValues.OutsideEnd) {
                return OfficeChartDataLabelPosition.OutsideEnd;
            }

            if (position == DataLabelPositionValues.Left) {
                return OfficeChartDataLabelPosition.Left;
            }

            if (position == DataLabelPositionValues.Right) {
                return OfficeChartDataLabelPosition.Right;
            }

            if (position == DataLabelPositionValues.Top) {
                return OfficeChartDataLabelPosition.Top;
            }

            if (position == DataLabelPositionValues.Bottom) {
                return OfficeChartDataLabelPosition.Bottom;
            }

            return OfficeChartDataLabelPosition.Auto;
        }

        private static bool HasNativeWordChartLegend(Chart chart) =>
            chart.GetFirstChild<Legend>() != null;

        private static OfficeChartLegendPosition GetNativeWordChartLegendPosition(Chart chart) {
            LegendPosition? position = chart.GetFirstChild<Legend>()?.GetFirstChild<LegendPosition>();
            if (position?.Val?.Value == LegendPositionValues.Left) {
                return OfficeChartLegendPosition.Left;
            }

            if (position?.Val?.Value == LegendPositionValues.Top) {
                return OfficeChartLegendPosition.Top;
            }

            if (position?.Val?.Value == LegendPositionValues.Bottom) {
                return OfficeChartLegendPosition.Bottom;
            }

            return OfficeChartLegendPosition.Right;
        }

        private static int? GetNativeWordChartMaximumCategoryAxisLabels(OpenXmlElement chartElement, PlotArea plotArea, int categoryCount) {
            if (categoryCount <= 0) {
                return null;
            }

            uint? skip = GetNativeWordChartCategoryAxisTickLabelSkip(chartElement, plotArea);
            if (!skip.HasValue || skip.Value <= 1U) {
                return null;
            }

            return Math.Max(1, (int)Math.Ceiling(categoryCount / (double)skip.Value));
        }

        private static uint? GetNativeWordChartCategoryAxisTickLabelSkip(OpenXmlElement chartElement, PlotArea plotArea) {
            var chartAxisIds = new HashSet<uint>(
                chartElement.Elements<AxisId>()
                    .Select(axis => axis.Val?.Value)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value));

            foreach (CategoryAxis axis in plotArea.Elements<CategoryAxis>()) {
                uint? axisId = axis.AxisId?.Val?.Value;
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    return GetNativeWordChartTickLabelSkip(axis);
                }
            }

            foreach (CategoryAxis axis in plotArea.Elements<CategoryAxis>()) {
                uint? skip = GetNativeWordChartTickLabelSkip(axis);
                if (skip.HasValue) {
                    return skip;
                }
            }

            return null;
        }

        private static uint? GetNativeWordChartTickLabelSkip(OpenXmlElement axis) {
            foreach (TickLabelSkip skip in axis.Descendants<TickLabelSkip>()) {
                var value = skip.Val?.Value;
                if (value.HasValue && value.Value > 1) {
                    return (uint)value.Value;
                }
            }

            return null;
        }

        private static bool IsNativeWordHorizontalBarChart(OfficeChartKind chartKind) =>
            chartKind == OfficeChartKind.BarClustered ||
            chartKind == OfficeChartKind.BarStacked ||
            chartKind == OfficeChartKind.BarStacked100;

        private sealed class NativeWordChartDataLabelState {
            public DataLabels? Labels { get; set; }
            public IReadOnlyCollection<int>? SeriesIndexes { get; set; }
            public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? PointIndexes { get; set; }
            public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? HiddenPointIndexes { get; set; }
            public bool ShowValue { get; set; }
            public bool ShowPercent { get; set; }
            public bool ShowCategoryName { get; set; }
            public bool ShowSeriesName { get; set; }
        }

        private static NativeWordChartDataLabelState GetNativeWordChartDataLabels(OpenXmlElement chartElement, OfficeChartKind chartKind) {
            var state = new NativeWordChartDataLabelState();
            List<int> renderableSeriesIndexes = GetNativeWordRenderableSeriesIndexes(chartElement, chartKind);
            DataLabels? chartLabels = chartElement.GetFirstChild<DataLabels>();
            if (chartLabels != null) {
                state.Labels = chartLabels;
                state.ShowValue = HasNativeWordChartDataLabelFlag<ShowValue>(chartLabels);
                state.ShowPercent = HasNativeWordChartDataLabelFlag<ShowPercent>(chartLabels);
                state.ShowCategoryName = HasNativeWordChartDataLabelFlag<ShowCategoryName>(chartLabels);
                state.ShowSeriesName = HasNativeWordChartDataLabelFlag<ShowSeriesName>(chartLabels);
                bool chartWideLabels = HasNativeWordChartDataLabelFlags(chartLabels);
                Dictionary<int, IReadOnlyCollection<int>>? visiblePoints = GetNativeWordChartDataLabelPointIndexes(chartLabels, renderableSeriesIndexes, deleted: false);
                Dictionary<int, IReadOnlyCollection<int>>? hiddenPoints = GetNativeWordChartDataLabelPointIndexes(chartLabels, renderableSeriesIndexes, deleted: true);
                state.PointIndexes = chartWideLabels ? null : visiblePoints;
                state.HiddenPointIndexes = hiddenPoints;
                return state;
            }

            var seriesIndexes = new List<int>();
            var pointIndexes = new Dictionary<int, IReadOnlyCollection<int>>();
            var hiddenPointIndexes = new Dictionary<int, IReadOnlyCollection<int>>();
            int seriesIndex = 0;
            foreach ((OpenXmlElement seriesElement, _) in GetNativeWordOrderedSeriesElements(chartElement)) {
                if (!IsNativeWordChartSeriesRenderable(seriesElement, chartKind)) {
                    continue;
                }

                DataLabels? labels = seriesElement.GetFirstChild<DataLabels>();
                if (labels != null) {
                    if (HasNativeWordChartDataLabelFlags(labels)) {
                        state.Labels = labels;
                        state.ShowValue |= HasNativeWordChartDataLabelFlag<ShowValue>(labels);
                        state.ShowPercent |= HasNativeWordChartDataLabelFlag<ShowPercent>(labels);
                        state.ShowCategoryName |= HasNativeWordChartDataLabelFlag<ShowCategoryName>(labels);
                        state.ShowSeriesName |= HasNativeWordChartDataLabelFlag<ShowSeriesName>(labels);
                        seriesIndexes.Add(seriesIndex);
                    } else {
                        state.Labels ??= labels;
                    }

                    AddNativeWordChartDataLabelPointIndexes(labels, seriesIndex, pointIndexes, deleted: false);
                    AddNativeWordChartDataLabelPointIndexes(labels, seriesIndex, hiddenPointIndexes, deleted: true);
                }

                seriesIndex++;
            }

            state.SeriesIndexes = seriesIndexes.Count == 0 ? null : seriesIndexes;
            state.PointIndexes = pointIndexes.Count == 0 ? null : pointIndexes;
            state.HiddenPointIndexes = hiddenPointIndexes.Count == 0 ? null : hiddenPointIndexes;
            return state;
        }

        private static List<int> GetNativeWordRenderableSeriesIndexes(OpenXmlElement chartElement, OfficeChartKind chartKind) {
            var indexes = new List<int>();
            int seriesIndex = 0;
            foreach (OpenXmlElement seriesElement in chartElement.ChildElements.Where(element => element.LocalName == "ser")) {
                if (IsNativeWordChartSeriesRenderable(seriesElement, chartKind)) {
                    if (seriesIndex >= MaxNativeWordChartSeries) {
                        throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
                    }

                    indexes.Add(seriesIndex);
                    seriesIndex++;
                }
            }

            return indexes;
        }

        private static Dictionary<int, IReadOnlyCollection<int>>? GetNativeWordChartDataLabelPointIndexes(DataLabels labels, IReadOnlyList<int> seriesIndexes, bool deleted) {
            var indexes = new Dictionary<int, IReadOnlyCollection<int>>();
            foreach (int seriesIndex in seriesIndexes) {
                AddNativeWordChartDataLabelPointIndexes(labels, seriesIndex, indexes, deleted);
            }

            return indexes.Count == 0 ? null : indexes;
        }

        private static void AddNativeWordChartDataLabelPointIndexes(DataLabels labels, int seriesIndex, Dictionary<int, IReadOnlyCollection<int>> indexes, bool deleted) {
            var pointIndexes = new List<int>();
            foreach (DataLabel label in labels.Elements<DataLabel>()) {
                uint? index = label.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>()?.Val?.Value;
                if (!index.HasValue) {
                    continue;
                }

                EnsureNativeWordChartPointIndex(index.Value);
                bool isDeleted = IsNativeWordChartBooleanOn(label.GetFirstChild<Delete>());
                if (deleted == isDeleted && (deleted || HasNativeWordChartDataLabelFlags(label))) {
                    pointIndexes.Add((int)index.Value);
                }
            }

            if (pointIndexes.Count > 0) {
                indexes[seriesIndex] = pointIndexes;
            }
        }

        private static bool HasNativeWordChartDataLabelFlags(OpenXmlElement labels) =>
            HasNativeWordChartDataLabelFlag<ShowValue>(labels) ||
            HasNativeWordChartDataLabelFlag<ShowPercent>(labels) ||
            HasNativeWordChartDataLabelFlag<ShowCategoryName>(labels) ||
            HasNativeWordChartDataLabelFlag<ShowSeriesName>(labels);

        private static bool HasNativeWordChartDataLabelFlag<T>(OpenXmlElement? labels)
            where T : BooleanType {
            if (labels == null) {
                return false;
            }

            T? value = labels.GetFirstChild<T>();
            if (IsNativeWordChartBooleanOn(value)) {
                return true;
            }

            return labels.Elements<DataLabel>().Any(label => IsNativeWordChartBooleanOn(label.GetFirstChild<T>()));
        }

        private static bool IsNativeWordChartLegendOverlay(Chart chart) =>
            IsNativeWordChartBooleanOn(chart.GetFirstChild<Legend>()?.GetFirstChild<Overlay>());

        private static bool IsNativeWordChartTitleOverlay(Chart chart) =>
            IsNativeWordChartBooleanOn(chart.Title?.GetFirstChild<Overlay>());

        private static bool IsNativeWordChartBooleanOn(BooleanType? value) =>
            value != null && (value.Val == null || value.Val.Value);

        private static void EnsureNativeWordChartPointIndex(uint index) {
            if (index >= MaxNativeWordChartPoints) {
                throw new NativeWordChartLimitException("Word chart cache contains a point index beyond the maximum supported point count.");
            }
        }

        private static void EnsureNativeWordChartPointCount(int count) {
            if (count > MaxNativeWordChartPoints) {
                throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported point count.");
            }
        }

        private sealed class NativeWordChartLimitException : Exception {
            public NativeWordChartLimitException(string message) : base(message) {
            }
        }
    }
}
