using System.Globalization;
using OfficeIMO.Drawing;
using static OfficeIMO.Word.OpenXml.WordXml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const double NativeEmusPerPoint = 12700D;
        private const int MaxNativeWordChartSeries = 256;
        private const int MaxNativeWordChartPoints = 4096;
        // The chart fills its wp:extent frame. The bounds only reject degenerate or runaway
        // frames; page-width fitting happens when the drawing is placed (NativeChartDrawingStyle).
        private const double MinNativeWordChartWidthPoints = 72D;
        private const double MaxNativeWordChartWidthPoints = 1584D;
        private const double MinNativeWordChartHeightPoints = 54D;
        // Fits the text column of a landscape Letter page with one-inch margins.
        private const double MaxNativeWordChartHeightPoints = 460D;
        private const double NativeWordChartTitleTopPadding = 31D;
        /// <summary>Places a chart at its frame size, scaled as a whole when the frame is wider than the text column.</summary>
        private static PdfCore.PdfDrawingStyle NativeChartDrawingStyle() => new() { ConstrainToContentWidth = true };

        private static OfficeDrawing? PrepareNativeChart(WordChart? chart, WordPdfExportOptions? options, string source) {
            if (chart == null) {
                return null;
            }

            OfficeChartSnapshot? snapshot;
            string? warning;
            IReadOnlyList<string> approximations;
            using (var fidelity = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Enter(options?.ChartRendering.Fidelity ?? OfficeChartFidelity.Strict)) {
                bool projected = TryCreateNativeWordChartSnapshot(chart, out snapshot, out warning);
                approximations = projected ? fidelity.Approximations.ToArray() : Array.Empty<string>();
                if (!projected) snapshot = null;
            }
            if (snapshot == null) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "native-body-chart-unsupported",
                        source,
                        warning ?? "Word chart data is not mapped by the OfficeIMO PDF engine yet.");
                }

                return CreateNativeChartPlaceholder(chart);
            }

            if (options != null) {
                foreach (string approximation in approximations)
                    AddNativeExportWarning(options, OfficeDiagnosticCodes.ChartApproximation, source, approximation);
                OfficeXmlNode nativePlotArea = GetNativeWordChartPlotArea(GetNativeWordChartElement(chart.ChartSpaceNode()));
                bool emittedProjection = false;
                if (nativePlotArea.Exists) {
                    foreach (OfficeXmlNode group in nativePlotArea.Elements()) {
                        string? projection = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.DescribeProjection(group);
                        if (projection == null) continue;
                        bool flattened = group.LocalName.EndsWith("3DChart", StringComparison.Ordinal);
                        AddNativeExportWarning(options,
                            flattened ? OfficeDiagnosticCodes.ChartFlattened : OfficeDiagnosticCodes.NativeBodyChartSimplified,
                            source, projection);
                        emittedProjection = true;
                    }
                }
                if (!emittedProjection && !string.IsNullOrWhiteSpace(warning))
                    AddNativeExportWarning(options, OfficeDiagnosticCodes.NativeBodyChartSimplified, source, warning!);
            }

            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(snapshot!);
            if (rendering.QualityReport.HasIssues && options != null) {
                AddNativeExportWarning(
                    options,
                    "chart-quality",
                    source,
                    "Exported Word chart '" + GetNativeWordChartDisplayName(snapshot!) + "' with shared drawing quality warnings: " + string.Join("; ", rendering.QualityReport.Issues.Select(issue => issue.ToString())));
            }

            return rendering.Drawing;
        }

        /// <summary>
        /// A chart the renderer cannot draw keeps its frame on the page as a labelled box,
        /// so the layout matches the document and the reader can see a chart belongs there.
        /// </summary>
        private static OfficeDrawing CreateNativeChartPlaceholder(WordChart chart) {
            (double width, double height) = GetNativeWordChartSizePoints(chart);
            var drawing = new OfficeDrawing(width, height);
            OfficeShape frame = OfficeShape.RoundedRectangle(width - 1D, height - 1D, 4D);
            frame.FillColor = OfficeColor.FromRgb(0xF6, 0xF8, 0xFB);
            frame.StrokeColor = OfficeColor.FromRgb(0xC9, 0xD2, 0xDE);
            frame.StrokeWidth = 0.75D;
            drawing.AddShape(frame, 0.5D, 0.5D);
            string? title = GetNativeWordChartElement(chart.ChartSpaceNode()) is { IsNone: false } openXmlChart
                ? GetNativeWordChartTitle(openXmlChart)
                : null;
            string label = string.IsNullOrWhiteSpace(title) ? "Chart" : title!;
            drawing.AddText(label, 12D, height / 2D - 20D, width - 24D, 18D,
                new OfficeFontInfo("Arial", 12D, OfficeFontStyle.Bold), OfficeColor.FromRgb(0x1B, 0x24, 0x33),
                OfficeTextAlignment.Center, verticalAlignment: OfficeTextVerticalAlignment.Center);
            drawing.AddText("This chart type is not drawn in the PDF export.", 12D, height / 2D + 2D, width - 24D, 16D,
                new OfficeFontInfo("Arial", 9D), OfficeColor.FromRgb(0x5B, 0x67, 0x77),
                OfficeTextAlignment.Center, verticalAlignment: OfficeTextVerticalAlignment.Center);
            return drawing;
        }

        internal static bool TryCreateNativeWordChartSnapshot(WordChart chart, out OfficeChartSnapshot? snapshot, out string? warning) {
            snapshot = null;
            warning = null;

            OfficeXmlNode chartSpace = chart.ChartSpaceNode();
            OfficeXmlNode openXmlChart = GetNativeWordChartElement(chartSpace);
            OfficeXmlNode plotArea = GetNativeWordChartPlotArea(openXmlChart);
            if (plotArea.IsNone) {
                warning = "Word chart part does not contain a plot area with cached chart data.";
                return false;
            }

            List<OfficeXmlNode> allChartElements = plotArea.Elements()
                .Where(IsNativeWordChartElement)
                .ToList();
            // The shared reader owns qualified legend-frame appearance, including its native
            // fill and outline. Keep the PDF chart on that same projection when it is styled. Error bars and
            // trendlines on a flat chart without a data table go to the shared reader too, which draws or reports
            // them; 3-D charts and data tables keep this path, which reports them below.
            bool seriesExtras = plotArea.DescendantsSkippingExtLst().Any(element => element.Is(C, "errBars") || element.Is(C, "trendline"));
            bool sharedReadsExtras = plotArea.Child(C, "dTable").IsNone &&
                !allChartElements.Any(element => element.LocalName.EndsWith("3DChart", StringComparison.Ordinal));
            if (allChartElements.Count > 1 || allChartElements.Any(element => element.Is(C, "bubbleChart")) ||
                seriesExtras && sharedReadsExtras ||
                HasMaterialNativeLegendFrame(openXmlChart) ||
                (plotArea.Child(C, "dTable").IsNone &&
                 (HasNativeExplicitValueAxisScale(plotArea) || HasMaterialNativeChartAreaShape(chartSpace)))) {
                return TryCreateSharedNativeWordChartSnapshot(chart, out snapshot, out warning);
            }
            if (seriesExtras) {
                if (!OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                    warning = "Error bars and trendlines on 3-D charts or charts with a data table cannot be projected.";
                    return false;
                }
                if (plotArea.HasDescendantSkippingExtLst(C, "errBars"))
                    OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Record("Error bars are left out.");
                if (plotArea.HasDescendantSkippingExtLst(C, "trendline"))
                    OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Record("Trendlines are left out.");
            }
            List<OfficeXmlNode> chartElements = allChartElements
                .Where(IsNativeSupportedWordChartElement)
                .ToList();
            if (chartElements.Count == 0) {
                warning = "Word chart type is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            OfficeXmlNode chartElement = chartElements[0];
            if (!TryMapNativeWordChartKind(chartElement, out OfficeChartType chartKind)) {
                warning = "Word chart type '" + chartElement.LocalName + "' is not supported by the shared OfficeIMO chart renderer.";
                return false;
            }

            try {
                IReadOnlyDictionary<string, OfficeColor> themeColors = GetNativeDrawingThemeColors(chartSpace);
                OfficeXmlNode pointColorScheme = GetNativeWordChartColorScheme(chartSpace);
                IReadOnlyList<OfficeSeries> series = ExtractNativeWordChartSeries(openXmlChart, chartElement, chartKind, themeColors, pointColorScheme, out IReadOnlyList<string> categories);
                if (categories.Count == 0 || series.Count == 0) {
                    warning = "Word chart does not contain cached categories and values that can be rendered without Office.";
                    return false;
                }

                (double width, double height) = GetNativeWordChartSizePoints(chart);
                string? title = GetNativeWordChartTitle(openXmlChart);
                string name = GetNativeWordChartName(chart, title);
                OfficeChartStyle? style = CreateNativeWordChartStyle(openXmlChart, chartElement, plotArea, chartKind, categories.Count, series.Count, themeColors, pointColorScheme);
                OfficeChartLayout? layout = CreateNativeWordChartLayout(openXmlChart, chartElement, plotArea, chartKind, categories.Count);
                OfficeChartRadialLayout radialLayout;
                try {
                    radialLayout = chart.RadialLayout;
                } catch (Exception exception) when (exception is ArgumentOutOfRangeException || exception is FormatException || exception is OverflowException) {
                    warning = "Word chart contains invalid pie rotation or doughnut hole metadata and cannot be rendered.";
                    return false;
                }
                snapshot = new OfficeChartSnapshot(
                    name,
                    title,
                    chartKind,
                    new OfficeChartData(categories, series),
                    global::OfficeIMO.OfficeLength.Points(width),
                    global::OfficeIMO.OfficeLength.Points(height),
                    style: style,
                    layout: layout,
                    radialLayout: radialLayout);
            } catch (NativeWordChartLimitException ex) {
                warning = ex.Message;
                return false;
            }

            warning = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.DescribeProjection(chartElement);
            return true;
        }

        /// <summary>The <c>c:chart</c> of a chart part's <c>c:chartSpace</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeWordChartElement(OfficeXmlNode chartSpace) =>
            chartSpace.IsNone ? default : chartSpace.Child(C, "chart");

        /// <summary>The <c>c:plotArea</c> of a <c>c:chart</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeWordChartPlotArea(OfficeXmlNode chart) =>
            chart.IsNone ? default : chart.Child(C, "plotArea");

        /// <summary>The theme colour scheme of the document that holds the chart part, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeWordChartColorScheme(OfficeXmlNode chartSpace) {
            OfficeXmlNode theme = chartSpace.IsNone ? default : chartSpace.PackageThemeNode();
            OfficeXmlNode themeElements = theme.IsNone ? default : theme.Child(A, "themeElements");
            return themeElements.IsNone ? default : themeElements.Child(A, "clrScheme");
        }

        /// <summary>The first <c>c:</c> child with this name, or <see cref="OfficeXmlNode.None"/> when the element is missing.</summary>
        private static OfficeXmlNode GetNativeWordChartChild(OfficeXmlNode element, string localName) =>
            element.IsNone ? default : element.Child(C, localName);

        private static bool HasMaterialNativeLegendFrame(OfficeXmlNode chart) {
            OfficeXmlNode properties = GetNativeWordChartChild(GetNativeWordChartChild(chart, "legend"), "spPr");
            return properties.Exists && properties.Elements().Any(child => child.Is(A, "ln")
                ? child.Child(A, "noFill").IsNone
                : !child.Is(A, "noFill") && (child.HasAttributes || child.HasChildren()));
        }

        private static bool IsNativeSupportedWordChartElement(OfficeXmlNode element) =>
            element.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or "stockChart" or "areaChart" or "area3DChart" or "pieChart" or "pie3DChart" or "doughnutChart" or "scatterChart" or "radarChart";

        private static bool IsNativeWordChartElement(OfficeXmlNode element) =>
            element.LocalName.EndsWith("Chart", StringComparison.Ordinal);

        private static bool TryMapNativeWordChartKind(OfficeXmlNode chartElement, out OfficeChartType kind) {
            switch (chartElement.LocalName) {
                case "barChart":
                case "bar3DChart":
                    bool horizontal = chartElement.Child(C, "barDir").Attr("val") == "bar";
                    string grouping = chartElement.Child(C, "grouping").Attr("val") ?? string.Empty;
                    kind = MapNativeWordBarChartKind(horizontal, grouping);
                    return true;
                case "stockChart":
                    kind = OfficeChartType.Line;
                    return true;
                case "lineChart":
                case "line3DChart":
                    kind = MapNativeWordLineChartKind(chartElement.Child(C, "grouping").Attr("val") ?? string.Empty);
                    return true;
                case "areaChart":
                case "area3DChart":
                    kind = MapNativeWordAreaChartKind(chartElement.Child(C, "grouping").Attr("val") ?? string.Empty);
                    return true;
                case "pieChart":
                case "pie3DChart":
                    kind = OfficeChartType.Pie;
                    return true;
                case "doughnutChart":
                    kind = OfficeChartType.Doughnut;
                    return true;
                case "scatterChart":
                    kind = OfficeChartType.Scatter;
                    return true;
                case "radarChart":
                    kind = OfficeChartType.Radar;
                    return true;
                default:
                    kind = default;
                    return false;
            }
        }

        private static OfficeChartType MapNativeWordBarChartKind(bool horizontal, string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return horizontal ? OfficeChartType.BarStacked : OfficeChartType.ColumnStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return horizontal ? OfficeChartType.BarStacked100 : OfficeChartType.ColumnStacked100;
            }

            return horizontal ? OfficeChartType.BarClustered : OfficeChartType.ColumnClustered;
        }

        private static OfficeChartType MapNativeWordLineChartKind(string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartType.LineStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartType.LineStacked100;
            }

            return OfficeChartType.Line;
        }

        private static OfficeChartType MapNativeWordAreaChartKind(string grouping) {
            if (grouping.Equals("Stacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartType.AreaStacked;
            }

            if (grouping.Equals("PercentStacked", StringComparison.OrdinalIgnoreCase)) {
                return OfficeChartType.AreaStacked100;
            }

            return OfficeChartType.Area;
        }

        internal static IReadOnlyList<OfficeSeries> ExtractNativeWordChartSeries(OfficeXmlNode chart, OfficeXmlNode chartElement, OfficeChartType chartKind, IReadOnlyDictionary<string, OfficeColor> themeColors, OfficeXmlNode pointColorScheme, out IReadOnlyList<string> categories) {
            var series = new List<OfficeSeries>();
            var categoryList = new List<string>();
            bool isScatter = chartKind == OfficeChartType.Scatter;
            bool varyColorsByPoint = !isScatter && !IsNativeWordPieLikeChart(chartKind) && IsNativeWordChartVaryColorsEnabled(chartElement);
            HashSet<uint> hiddenLegendIndexes = GetNativeWordHiddenLegendIndexes(chart);
            var automaticColors = new OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext();

            int seriesIndex = 0;
            foreach ((OfficeXmlNode seriesElement, int originalSeriesIndex) in GetNativeWordOrderedSeriesElements(chartElement)) {
                if (seriesIndex >= MaxNativeWordChartSeries) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
                }

                IReadOnlyList<double> values;
                IReadOnlyList<double>? xValues = null;
                IReadOnlyList<string> currentCategories;
                if (isScatter) {
                    xValues = ExtractNativeWordChartNumberValues(seriesElement.Child(C, "xVal"));
                    values = ExtractNativeWordChartNumberValues(seriesElement.Child(C, "yVal"));
                    if (xValues.Count != values.Count) {
                        int count = Math.Max(xValues.Count, values.Count);
                        xValues = NormalizeNativeWordChartNumberValues(xValues, count, useIndexDefaults: true);
                        values = NormalizeNativeWordChartNumberValues(values, count, useIndexDefaults: false);
                    }

                    currentCategories = xValues.Count > 0
                        ? xValues.Select(value => value.ToString("0.####", CultureInfo.InvariantCulture)).ToList()
                        : CreateNativeWordChartDefaultCategories(values.Count);
                } else {
                    values = ExtractNativeWordChartNumberValues(seriesElement.Child(C, "val"));
                    currentCategories = ExtractNativeWordChartCategories(seriesElement.Child(C, "cat"), values.Count);
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

                IReadOnlyList<OfficeXmlNode> pointOverrides = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPointStyles
                    .GetBoundedPoints(seriesElement);
                IReadOnlyList<OfficeColor?>? pointColors = ExtractNativeWordChartPointColors(pointOverrides, values.Count, themeColors);
                if (pointColors == null && varyColorsByPoint && seriesIndex == 0) {
                    pointColors = CreateNativeWordChartVaryPointColors(chart, values.Count, pointColorScheme);
                }

                OfficeColor? seriesColor = null;
                if (IsNativeWordPieLikeChart(chartKind) &&
                    TryGetNativeWordChartFillColor(seriesElement, themeColors, out OfficeColor pieLikeSeriesColor)) {
                    seriesColor = pieLikeSeriesColor;
                }

                bool showMarkers = !IsNativeWordChartSeriesMarkerHidden(seriesElement);
                bool connectLine = !IsNativeWordLineLikeChart(chartKind) || !HasNativeDrawingOutlineNoFill(seriesElement.Child(C, "spPr"));
                var seriesPaint = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPointStyles.ReadSeriesStyle(seriesElement, pointColorScheme);
                if (seriesElement.Parent.Is(C, "line3DChart")) {
                    var ribbon = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPointStyles.ReadRibbonPaint(seriesElement, pointColorScheme, automaticColors);
                    seriesPaint = ribbon.Paint;
                    seriesColor = ribbon.Color;
                    showMarkers = false;
                    connectLine = ribbon.Paint.ShowOutline != false;
                }
                OfficeSeries projected = new OfficeSeries(
                    GetNativeWordChartSeriesName(seriesElement, seriesIndex),
                    values,
                    xValues,
                    seriesColor,
                    pointColors,
                    showMarkers,
                    !hiddenLegendIndexes.Contains((uint)originalSeriesIndex),
                    connectLine,
                    markerSize: GetNativeWordChartMarkerSize(seriesElement, chartElement, chartKind),
                    markerShape: GetNativeWordChartMarkerShape(seriesElement, chartElement, chartKind))
                    .WithSeriesStyle(seriesPaint)
                    .WithPointStyles(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartPointStyles.Read(pointOverrides, values.Count,
                        pointColorScheme, seriesElement.Parent.LocalName.EndsWith("3DChart", StringComparison.Ordinal)));
                if (IsNativeWordPieLikeChart(chartKind)) {
                    if (!OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartExplosions.TryRead(
                            seriesElement, pointOverrides, values.Count,
                            out int[]? explosions))
                        throw new NativeWordChartLimitException("Word chart contains unsupported slice explosion metadata.");
                    projected = projected.WithPointExplosions(explosions);
                }
                series.Add(projected);
                seriesIndex++;
            }

            categories = categoryList;
            return series;
        }

        private static IReadOnlyList<string> ExtractNativeWordChartCategories(OfficeXmlNode categoryAxisData, int valueCount) {
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

        private static bool IsNativeWordChartVaryColorsEnabled(OfficeXmlNode chartElement) =>
            IsNativeWordChartBooleanOn(chartElement.Child(C, "varyColors"));

        private static bool IsNativeWordChartVaryColorsDisabled(OfficeXmlNode chartElement) {
            string? value = chartElement.Child(C, "varyColors").Attr("val");
            return value != null && ParseNativeWordChartBoolean(value) == false;
        }

        private static HashSet<uint> GetNativeWordHiddenLegendIndexes(OfficeXmlNode chart) {
            var indexes = new HashSet<uint>();
            OfficeXmlNode legend = chart.Child(C, "legend");
            if (legend.IsNone) {
                return indexes;
            }

            foreach (OfficeXmlNode entry in legend.Children(C, "legendEntry")) {
                uint? index = GetNativeWordChartIndex(entry);
                if (index.HasValue && IsNativeWordChartBooleanOn(entry.Child(C, "delete"))) {
                    indexes.Add(index.Value);
                }
            }

            return indexes;
        }

        private static IReadOnlyList<OfficeColor?> CreateNativeWordChartVaryPointColors(OfficeXmlNode chart, int valueCount, OfficeXmlNode scheme) {
            OfficeColor[]? automatic = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ResolvePoints(chart.Parent, valueCount, scheme);
            var colors = new OfficeColor?[valueCount];
            for (int index = 0; index < colors.Length; index++) {
                colors[index] = automatic?[index] ?? OfficeChartStyle.Default.GetSeriesColor(index);
            }

            return colors;
        }

        private static IReadOnlyList<string> ExtractNativeWordChartStringValues(OfficeXmlNode container) {
            if (container.IsNone) {
                return Array.Empty<string>();
            }

            var values = new SortedDictionary<uint, string>();
            uint fallbackIndex = 0;
            uint maxIndex = 0;
            bool hasPoint = false;
            foreach (OfficeXmlNode point in container.DescendantsSkippingExtLst(C, "pt").Where(IsNativeWordChartStringPoint)) {
                if (fallbackIndex >= MaxNativeWordChartPoints) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported point count.");
                }

                uint index = ParseUInt32(point.Attr("idx")) ?? fallbackIndex;
                EnsureNativeWordChartPointIndex(index);
                hasPoint = true;
                if (index > maxIndex) {
                    maxIndex = index;
                }

                string? value = GetNativeWordChartPointText(point);
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

        private static IReadOnlyList<double> ExtractNativeWordChartNumberValues(OfficeXmlNode container) {
            if (container.IsNone) {
                return Array.Empty<double>();
            }

            var values = new SortedDictionary<uint, double>();
            uint fallbackIndex = 0;
            uint maxIndex = 0;
            bool hasPoint = false;
            foreach (OfficeXmlNode point in container.DescendantsSkippingExtLst(C, "pt").Where(IsNativeWordChartNumericPoint)) {
                if (fallbackIndex >= MaxNativeWordChartPoints) {
                    throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported point count.");
                }

                uint index = ParseUInt32(point.Attr("idx")) ?? fallbackIndex;
                EnsureNativeWordChartPointIndex(index);
                hasPoint = true;
                if (index > maxIndex) {
                    maxIndex = index;
                }

                if (TryParseNativeWordChartNumber(GetNativeWordChartPointText(point), out double value)) {
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

        // c:pt is a string point in a string cache or literal (and a multi-level level), a numeric point in a
        // numeric cache or literal; the readers take each kind only from its own containers.
        private static bool IsNativeWordChartStringPoint(OfficeXmlNode point) {
            OfficeXmlNode parent = point.Parent;
            return parent.Is(C, "strCache") || parent.Is(C, "strLit") || parent.Is(C, "lvl");
        }

        private static bool IsNativeWordChartNumericPoint(OfficeXmlNode point) {
            OfficeXmlNode parent = point.Parent;
            return parent.Is(C, "numCache") || parent.Is(C, "numLit");
        }

        /// <summary>The text of a point's <c>c:v</c>, or null when it has none.</summary>
        private static string? GetNativeWordChartPointText(OfficeXmlNode point) {
            OfficeXmlNode value = point.Child(C, "v");
            return value.IsNone ? null : value.InnerText;
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
            if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                if (!double.IsNaN(value) && !double.IsInfinity(value)) {
                    return true;
                }
            }

            value = 0D;
            return false;
        }

        private static IReadOnlyList<(OfficeXmlNode SeriesElement, int OriginalSeriesIndex)> GetNativeWordOrderedSeriesElements(OfficeXmlNode chartElement) {
            var series = new List<(OfficeXmlNode SeriesElement, int OriginalSeriesIndex, int Order)>();
            int index = 0;
            foreach (OfficeXmlNode element in chartElement.Elements().Where(element => element.LocalName == "ser")) {
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

        private static int GetNativeWordChartSeriesOrder(OfficeXmlNode seriesElement, int fallback) {
            OfficeXmlNode order = seriesElement.Elements().FirstOrDefault(element => element.LocalName == "order");
            string? value = order.IsNone ? null : GetNativeOpenXmlAttribute(order, "val");
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : fallback;
        }

        private static string GetNativeWordChartSeriesName(OfficeXmlNode seriesElement, int index) {
            string? name = GetFirstNativeWordChartText(seriesElement.Child(C, "tx"));
            return string.IsNullOrWhiteSpace(name)
                ? "Series " + (index + 1).ToString(CultureInfo.InvariantCulture)
                : name!;
        }

        private static string? GetNativeWordChartTitle(OfficeXmlNode chart) =>
            GetFirstNativeWordChartText(GetNativeWordChartChild(chart, "title"));

        private static string? GetFirstNativeWordChartText(OfficeXmlNode element) {
            if (element.IsNone) {
                return null;
            }

            string drawingText = string.Concat(element.DescendantsSkippingExtLst(A, "t").Select(text => text.InnerText));
            if (!string.IsNullOrWhiteSpace(drawingText)) {
                return drawingText;
            }

            string cachedText = string.Concat(element.DescendantsSkippingExtLst(C, "v").Select(text => text.InnerText));
            return string.IsNullOrWhiteSpace(cachedText) ? null : cachedText;
        }

        private static string GetNativeWordChartName(WordChart chart, string? title) {
            OfficeXmlNode drawing = chart.DrawingNode();
            OfficeXmlNode inline = drawing.IsNone ? default : drawing.Child(Wp, "inline");
            string? drawingName = inline.IsNone ? null : inline.Child(Wp, "docPr").Attr("name");
            if (!string.IsNullOrWhiteSpace(drawingName)) {
                return drawingName!;
            }

            if (!string.IsNullOrWhiteSpace(title)) {
                return title!;
            }

            return chart.ChartPartUri() ?? "Word chart";
        }

        private static string GetNativeWordChartDisplayName(OfficeChartSnapshot snapshot) =>
            string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title!;

        private static (double Width, double Height) GetNativeWordChartSizePoints(WordChart chart) {
            OfficeXmlNode drawing = chart.DrawingNode();
            OfficeXmlNode inlineExtent = drawing.IsNone ? default : drawing.Child(Wp, "inline") is { IsNone: false } inline ? inline.Child(Wp, "extent") : default;
            OfficeXmlNode anchorExtent = drawing.IsNone ? default : drawing.Child(Wp, "anchor") is { IsNone: false } anchor ? anchor.Child(Wp, "extent") : default;
            long? cx = inlineExtent.Int64Attr("cx") ?? anchorExtent.Int64Attr("cx");
            long? cy = inlineExtent.Int64Attr("cy") ?? anchorExtent.Int64Attr("cy");
            double width = cx.HasValue && cx.Value > 0 ? cx.Value / NativeEmusPerPoint : 360D;
            double height = cy.HasValue && cy.Value > 0 ? cy.Value / NativeEmusPerPoint : 216D;
            // Scale an oversized frame as a whole so the chart keeps its proportions.
            double scale = Math.Min(1D, Math.Min(MaxNativeWordChartWidthPoints / width, MaxNativeWordChartHeightPoints / height));
            width = Math.Max(MinNativeWordChartWidthPoints, width * scale);
            height = Math.Max(MinNativeWordChartHeightPoints, height * scale);
            return (width, height);
        }

        private static OfficeChartStyle? CreateNativeWordChartStyle(OfficeXmlNode chart, OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind, int categoryCount, int seriesCount, IReadOnlyDictionary<string, OfficeColor> themeColors, OfficeXmlNode scheme) {
            EnsureNativeWordChartPointCount(categoryCount);
            if (seriesCount > MaxNativeWordChartSeries) {
                throw new NativeWordChartLimitException("Word chart cache exceeds the maximum supported series count.");
            }

            int paletteCount = IsNativeWordPieLikeChart(chartKind) ? categoryCount : seriesCount;
            if (paletteCount <= 0) {
                return null;
            }

            // Series and points without a stored colour take Office's automatic colours from the
            // chart style and theme; the OfficeIMO palette only covers an unresolvable theme.
            var palette = new List<OfficeColor>(paletteCount);
            OfficeColor[]? automaticPoints = IsNativeWordPieLikeChart(chartKind)
                ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ResolvePoints(chart.Parent, paletteCount, scheme)
                : null;
            OfficeXmlNode[] automaticSeries = GetNativeWordOrderedSeriesElements(chartElement)
                .Where(item => IsNativeWordChartSeriesRenderable(item.SeriesElement, chartKind))
                .Select(item => item.SeriesElement).ToArray();
            var automaticColors = new OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext();
            for (int index = 0; index < paletteCount; index++) {
                OfficeColor? automatic = automaticPoints?[index] ?? (index < automaticSeries.Length && !IsNativeWordPieLikeChart(chartKind)
                    ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ResolveSeries(
                        automaticSeries[index], scheme, automaticColors)
                    : null);
                palette.Add(automatic ?? OfficeChartStyle.Default.GetSeriesColor(index));
            }

            if (IsNativeWordPieLikeChart(chartKind)) {
                OfficeXmlNode seriesElement = chartElement.Elements().FirstOrDefault(element => element.LocalName == "ser" && IsNativeWordChartSeriesRenderable(element, chartKind));
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

                foreach (OfficeXmlNode point in seriesElement.IsNone ? Enumerable.Empty<OfficeXmlNode>() : seriesElement.Elements(C, "dPt")) {
                    uint? pointIndex = GetNativeWordChartIndex(point);
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
                foreach ((OfficeXmlNode seriesElement, _) in GetNativeWordOrderedSeriesElements(chartElement).Where(item => IsNativeWordChartSeriesRenderable(item.SeriesElement, chartKind))) {
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
            // The chart area's c:spPr belongs to c:chartSpace, the parent of c:chart; older
            // documents may carry it inside c:chart instead, which takes precedence.
            OfficeXmlNode chartShape = chart.Child(C, "spPr");
            if (chartShape.IsNone && chart.Parent.Exists) chartShape = chart.Parent.Child(C, "spPr");
            bool showBackground = !HasNativeDrawingNoFill(chartShape);
            bool hasExplicitChartNoFill = chartShape.Exists && !showBackground;
            bool showBorder = !HasNativeDrawingOutlineNoFill(chartShape);
            bool hasExplicitChartNoLine = chartShape.Exists && !showBorder;
            if (TryGetNativeDrawingSolidFillColor(chartShape, out OfficeColor chartFill, themeColors)) {
                backgroundColor = chartFill;
            }

            if (TryGetNativeDrawingOutlineColor(chartShape, out OfficeColor chartBorder, themeColors)) {
                borderColor = chartBorder;
            }

            OfficeColor? plotAreaBackgroundColor = null;
            OfficeColor? plotAreaBorderColor = null;
            OfficeXmlNode plotShape = plotArea.Child(C, "spPr");
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

        // A chart area with no fill and no outline, as OfficeIMO writes for new charts, is
        // drawn by the simple projection; any other chart-area appearance needs the shared reader.
        private static bool HasMaterialNativeChartAreaShape(OfficeXmlNode chartSpace) {
            OfficeXmlNode shape = GetNativeWordChartChild(chartSpace, "spPr");
            if (shape.IsNone) return false;
            foreach (OfficeXmlNode child in shape.Elements()) {
                if (child.Is(A, "noFill")) continue;
                if (child.Is(A, "ln") && child.Elements().All(item => item.Is(A, "noFill")) && !child.HasAttributes) continue;
                return true;
            }

            return false;
        }

        private static bool IsNativeWordChartSeriesRenderable(OfficeXmlNode seriesElement, OfficeChartType chartKind) {
            OfficeXmlNode valuesElement = chartKind == OfficeChartType.Scatter
                ? seriesElement.Child(C, "yVal")
                : seriesElement.Child(C, "val");
            return ExtractNativeWordChartNumberValues(valuesElement).Count > 0;
        }

        private static OfficeColor? GetNativeWordChartTitleColor(OfficeXmlNode chart, IReadOnlyDictionary<string, OfficeColor> themeColors) {
            OfficeXmlNode title = chart.Child(C, "title");
            if (title.IsNone) {
                return null;
            }

            foreach (OfficeXmlNode textProperties in title.DescendantsSkippingExtLst().Where(element => element.LocalName == "defRPr" || element.LocalName == "rPr")) {
                if (TryGetNativeDrawingSolidFillColor(textProperties, out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            return null;
        }

        private static OfficeColor? GetNativeWordChartAxisLineColor(OfficeXmlNode chartElement, OfficeXmlNode plotArea, IReadOnlyDictionary<string, OfficeColor> themeColors) {
            if (TryGetNativeWordChartAxisLineColor(NativeWordChartValueAxis, chartElement, plotArea, themeColors, out OfficeColor valueAxisColor)) {
                return valueAxisColor;
            }

            if (TryGetNativeWordChartAxisLineColor(NativeWordChartCategoryAxis, chartElement, plotArea, themeColors, out OfficeColor categoryAxisColor)) {
                return categoryAxisColor;
            }

            return null;
        }

        private const string NativeWordChartValueAxis = "valAx";
        private const string NativeWordChartCategoryAxis = "catAx";

        private static bool TryGetNativeWordChartAxisLineColor(string axisName, OfficeXmlNode chartElement, OfficeXmlNode plotArea, IReadOnlyDictionary<string, OfficeColor> themeColors, out OfficeColor color) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, axisName)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue &&
                    chartAxisIds.Contains(axisId.Value) &&
                    TryGetNativeDrawingOutlineColor(axis.Child(C, "spPr"), out color, themeColors)) {
                    return true;
                }
            }

            foreach (OfficeXmlNode axis in plotArea.Elements(C, axisName)) {
                if (TryGetNativeDrawingOutlineColor(axis.Child(C, "spPr"), out color, themeColors)) {
                    return true;
                }
            }

            color = default;
            return false;
        }

        private static bool HasNativeWordChartAxisLineNoFill(string axisName, OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, axisName)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    return HasNativeDrawingOutlineNoFill(axis.Child(C, "spPr"));
                }
            }

            OfficeXmlNode fallback = plotArea.Child(C, axisName);
            return HasNativeDrawingOutlineNoFill(GetNativeWordChartChild(fallback, "spPr"));
        }

        private static OfficeColor? GetNativeWordChartMajorGridLineColor(OfficeXmlNode chartElement, OfficeXmlNode plotArea, IReadOnlyDictionary<string, OfficeColor> themeColors) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue &&
                    chartAxisIds.Contains(axisId.Value) &&
                    TryGetNativeDrawingOutlineColor(GetNativeWordChartChild(axis.Child(C, "majorGridlines"), "spPr"), out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                if (TryGetNativeDrawingOutlineColor(GetNativeWordChartChild(axis.Child(C, "majorGridlines"), "spPr"), out OfficeColor color, themeColors)) {
                    return color;
                }
            }

            return null;
        }

        private static bool HasNativeWordChartMajorGridLines(OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            bool hasMatchingValueAxis = false;
            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    hasMatchingValueAxis = true;
                    OfficeXmlNode gridlines = axis.Child(C, "majorGridlines");
                    if (gridlines.Exists) {
                        return !HasNativeDrawingOutlineNoFill(gridlines.Child(C, "spPr"));
                    }
                }
            }

            if (hasMatchingValueAxis) {
                return false;
            }

            bool hasAnyValueAxis = false;
            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                hasAnyValueAxis = true;
                OfficeXmlNode gridlines = axis.Child(C, "majorGridlines");
                if (gridlines.Exists) {
                    return !HasNativeDrawingOutlineNoFill(gridlines.Child(C, "spPr"));
                }
            }

            return !hasAnyValueAxis;
        }

        private static bool IsNativeWordPieLikeChart(OfficeChartType chartKind) =>
            chartKind == OfficeChartType.Pie || chartKind == OfficeChartType.Doughnut;

        private static bool TryGetNativeWordChartFillColor(OfficeXmlNode element, IReadOnlyDictionary<string, OfficeColor> themeColors, out OfficeColor color) {
            return TryGetNativeDrawingSolidFillColor(GetNativeWordChartChild(element, "spPr"), out color, themeColors);
        }

        private static bool TryGetNativeWordChartSeriesColor(OfficeXmlNode element, OfficeChartType chartKind, IReadOnlyDictionary<string, OfficeColor> themeColors, out OfficeColor color) {
            OfficeXmlNode properties = GetNativeWordChartChild(element, "spPr");
            if (TryGetNativeDrawingSolidFillColor(properties, out color, themeColors)) {
                return true;
            }

            return IsNativeWordLineLikeChart(chartKind) && TryGetNativeDrawingOutlineColor(properties, out color, themeColors);
        }

        private static bool IsNativeWordLineLikeChart(OfficeChartType chartKind) =>
            chartKind == OfficeChartType.Line ||
            chartKind == OfficeChartType.LineStacked ||
            chartKind == OfficeChartType.LineStacked100 ||
            chartKind == OfficeChartType.Scatter ||
            chartKind == OfficeChartType.Radar;

        private static IReadOnlyList<OfficeColor?>? ExtractNativeWordChartPointColors(IReadOnlyList<OfficeXmlNode> points, int valueCount, IReadOnlyDictionary<string, OfficeColor> themeColors) {
            if (valueCount <= 0) {
                return null;
            }

            OfficeColor?[] colors = new OfficeColor?[valueCount];
            bool anyColor = false;
            foreach (OfficeXmlNode point in points) {
                uint? index = GetNativeWordChartIndex(point);
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

        private static OfficeChartLayout? CreateNativeWordChartLayout(OfficeXmlNode chart, OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind, int categoryCount) {
            NativeWordChartDataLabelState dataLabels = GetNativeWordChartDataLabels(chartElement, chartKind);
            OfficeXmlNode labels = dataLabels.Labels;
            bool showValue = dataLabels.ShowValue || HasNativeWordChartDataLabelFlag(NativeWordChartShowValue, labels);
            bool showPercent = labels.Exists && IsNativeWordPieLikeChart(chartKind) && (dataLabels.ShowPercent || HasNativeWordChartDataLabelFlag(NativeWordChartShowPercent, labels));
            bool showCategoryName = dataLabels.ShowCategoryName || HasNativeWordChartDataLabelFlag(NativeWordChartShowCategoryName, labels);
            bool showSeriesName = dataLabels.ShowSeriesName || HasNativeWordChartDataLabelFlag(NativeWordChartShowSeriesName, labels);
            bool showDataLabels = showValue || showPercent || showCategoryName || showSeriesName;
            bool showLegend = HasNativeWordChartLegend(chart);
            OfficeChartLegendPosition legendPosition = GetNativeOfficeChartLegendPosition(chart);
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
            bool fillRadarSeries = chartKind != OfficeChartType.Radar || IsNativeWordFilledRadar(chartElement, chartKind);
            bool showCategoryAxis = IsNativeWordChartCategoryAxisVisible(chartElement, plotArea, chartKind);
            bool showValueAxis = IsNativeWordChartValueAxisVisible(chartElement, plotArea, chartKind);
            bool showCategoryAxisLine = showCategoryAxis && !HasNativeWordChartAxisLineNoFill(NativeWordChartCategoryAxis, chartElement, plotArea);
            bool showValueAxisLine = showValueAxis && !HasNativeWordChartAxisLineNoFill(NativeWordChartValueAxis, chartElement, plotArea);
            bool showCategoryAxisLabels = showCategoryAxis && IsNativeWordChartCategoryAxisLabelsVisible(chartElement, plotArea, chartKind);
            bool showValueAxisLabels = showValueAxis && IsNativeWordChartValueAxisLabelsVisible(chartElement, plotArea, chartKind);
            bool overlayTitle = IsNativeWordChartTitleOverlay(chart);
            double? titleTopPadding = !overlayTitle && !string.IsNullOrWhiteSpace(GetNativeWordChartTitle(chart))
                ? NativeWordChartTitleTopPadding
                : null;
            if (chartKind == OfficeChartType.Scatter) {
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
                } else if (chartKind == OfficeChartType.Radar) {
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

            string? separator = labels.IsNone ? null : labels.Child(C, "separator") is { IsNone: false } separatorElement ? separatorElement.InnerText : null;
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

            return OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCategoryGeometry.WithCategoryGeometry(layout, plotArea);
        }

        private static bool IsNativeWordMarkerOnlyScatter(OfficeXmlNode chartElement, OfficeChartType chartKind) =>
            chartKind == OfficeChartType.Scatter &&
            chartElement.Child(C, "scatterStyle").Attr("val") == "marker";

        private static bool IsNativeWordFilledRadar(OfficeXmlNode chartElement, OfficeChartType chartKind) =>
            chartKind == OfficeChartType.Radar &&
            chartElement.Child(C, "radarStyle").Attr("val") == "filled";

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

        private static NativeWordScatterAxisMetadata GetNativeWordScatterAxisMetadata(OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            IReadOnlyList<uint> valueAxisIds = GetNativeWordChartAxisIds(chartElement)
                .Where(axisId => GetNativeWordChartValueAxis(plotArea, axisId).Exists)
                .ToArray();

            OfficeXmlNode horizontalAxis = valueAxisIds.Count > 0 ? GetNativeWordChartValueAxis(plotArea, valueAxisIds[0]) : default;
            OfficeXmlNode verticalAxis = valueAxisIds.Count > 1 ? GetNativeWordChartValueAxis(plotArea, valueAxisIds[1]) : default;
            return new NativeWordScatterAxisMetadata(
                GetNativeWordChartAxisNumberFormat(horizontalAxis),
                GetNativeWordChartAxisNumberFormat(verticalAxis),
                horizontalAxis.IsNone ? null : GetNativeWordChartAxisTitle(horizontalAxis),
                verticalAxis.IsNone ? null : GetNativeWordChartAxisTitle(verticalAxis),
                !IsNativeWordChartAxisDeleted(horizontalAxis),
                !IsNativeWordChartAxisDeleted(verticalAxis),
                !IsNativeWordChartAxisDeleted(horizontalAxis) && !HasNativeDrawingOutlineNoFill(GetNativeWordChartChild(horizontalAxis, "spPr")),
                !IsNativeWordChartAxisDeleted(verticalAxis) && !HasNativeDrawingOutlineNoFill(GetNativeWordChartChild(verticalAxis, "spPr")),
                !IsNativeWordChartAxisDeleted(horizontalAxis) && !IsNativeWordChartTickLabelsHidden(horizontalAxis),
                !IsNativeWordChartAxisDeleted(verticalAxis) && !IsNativeWordChartTickLabelsHidden(verticalAxis));
        }

        private static string? GetNativeWordChartCategoryAxisTitle(OfficeXmlNode chartElement, OfficeXmlNode plotArea) =>
            GetNativeWordChartAxisTitle(NativeWordChartCategoryAxis, chartElement, plotArea);

        private static string? GetNativeWordChartValueAxisTitle(OfficeXmlNode chartElement, OfficeXmlNode plotArea) =>
            GetNativeWordChartAxisTitle(NativeWordChartValueAxis, chartElement, plotArea);

        private static string? GetNativeWordChartAxisTitle(string axisName, OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, axisName)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    string? title = GetNativeWordChartAxisTitle(axis);
                    if (!string.IsNullOrWhiteSpace(title)) {
                        return title;
                    }
                }
            }

            foreach (OfficeXmlNode axis in plotArea.Elements(C, axisName)) {
                string? title = GetNativeWordChartAxisTitle(axis);
                if (!string.IsNullOrWhiteSpace(title)) {
                    return title;
                }
            }

            return null;
        }

        private static string? GetNativeWordChartAxisTitle(OfficeXmlNode axis) =>
            GetFirstNativeWordChartText(axis.Child(C, "title"));

        private static IReadOnlyList<uint> GetNativeWordChartAxisIds(OfficeXmlNode chartElement) =>
            chartElement.Elements(C, "axId")
                .Select(axis => ParseUInt32(axis.Attr("val")))
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToArray();

        /// <summary>The <c>c:axId</c> value of an axis, or null.</summary>
        private static uint? GetNativeWordChartAxisId(OfficeXmlNode axis) => ParseUInt32(axis.Child(C, "axId").Attr("val"));

        /// <summary>The <c>c:idx</c> value of a series, point, label or legend entry, or null.</summary>
        private static uint? GetNativeWordChartIndex(OfficeXmlNode element) => ParseUInt32(element.Child(C, "idx").Attr("val"));

        private static OfficeXmlNode GetNativeWordChartValueAxis(OfficeXmlNode plotArea, uint axisId) =>
            plotArea.Elements(C, NativeWordChartValueAxis).FirstOrDefault(axis => GetNativeWordChartAxisId(axis) == axisId);

        private static OfficeXmlNode GetNativeWordChartCategoryAxis(OfficeXmlNode plotArea, uint axisId) =>
            plotArea.Elements(C, NativeWordChartCategoryAxis).FirstOrDefault(axis => GetNativeWordChartAxisId(axis) == axisId);

        private static bool IsNativeWordChartCategoryAxisVisible(OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind) {
            if (chartKind == OfficeChartType.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                OfficeXmlNode axis = GetNativeWordChartCategoryAxis(plotArea, axisId);
                if (axis.Exists) {
                    return !IsNativeWordChartAxisDeleted(axis);
                }
            }

            OfficeXmlNode fallback = plotArea.Child(C, NativeWordChartCategoryAxis);
            return !IsNativeWordChartAxisDeleted(fallback);
        }

        private static bool IsNativeWordChartValueAxisVisible(OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind) {
            if (chartKind == OfficeChartType.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                OfficeXmlNode axis = GetNativeWordChartValueAxis(plotArea, axisId);
                if (axis.Exists) {
                    return !IsNativeWordChartAxisDeleted(axis);
                }
            }

            OfficeXmlNode fallback = plotArea.Child(C, NativeWordChartValueAxis);
            return !IsNativeWordChartAxisDeleted(fallback);
        }

        private static bool IsNativeWordChartAxisDeleted(OfficeXmlNode axis) =>
            axis.Exists && IsNativeWordChartBooleanOn(axis.Child(C, "delete"));

        private static bool IsNativeWordChartCategoryAxisLabelsVisible(OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind) {
            if (chartKind == OfficeChartType.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                OfficeXmlNode axis = GetNativeWordChartCategoryAxis(plotArea, axisId);
                if (axis.Exists) {
                    return !IsNativeWordChartTickLabelsHidden(axis);
                }
            }

            return !IsNativeWordChartTickLabelsHidden(plotArea.Child(C, NativeWordChartCategoryAxis));
        }

        private static bool IsNativeWordChartValueAxisLabelsVisible(OfficeXmlNode chartElement, OfficeXmlNode plotArea, OfficeChartType chartKind) {
            if (chartKind == OfficeChartType.Radar || IsNativeWordPieLikeChart(chartKind)) {
                return true;
            }

            foreach (uint axisId in GetNativeWordChartAxisIds(chartElement)) {
                OfficeXmlNode axis = GetNativeWordChartValueAxis(plotArea, axisId);
                if (axis.Exists) {
                    return !IsNativeWordChartTickLabelsHidden(axis);
                }
            }

            return !IsNativeWordChartTickLabelsHidden(plotArea.Child(C, NativeWordChartValueAxis));
        }

        private static bool IsNativeWordChartTickLabelsHidden(OfficeXmlNode axis) {
            OfficeXmlNode position = axis.IsNone ? default : axis.Elements().FirstOrDefault(element => element.LocalName == "tickLblPos");
            return position.Exists && string.Equals(GetNativeOpenXmlAttribute(position, "val"), "none", StringComparison.OrdinalIgnoreCase);
        }

        private static string? GetNativeWordChartAxisNumberFormat(OfficeXmlNode axis) {
            string? format = GetNativeWordChartChild(axis, "numFmt").Attr("formatCode");
            return string.IsNullOrWhiteSpace(format) ? null : format;
        }

        private static string? GetNativeWordChartValueAxisNumberFormat(OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    string? format = GetNativeWordChartAxisNumberFormat(axis);
                    if (!string.IsNullOrWhiteSpace(format)) {
                        return format;
                    }
                }
            }

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartValueAxis)) {
                string? format = GetNativeWordChartAxisNumberFormat(axis);
                if (!string.IsNullOrWhiteSpace(format)) {
                    return format;
                }
            }

            return null;
        }

        private static bool AreNativeWordChartMarkersHidden(OfficeXmlNode chartElement, OfficeChartType chartKind) {
            if (chartKind != OfficeChartType.Line &&
                chartKind != OfficeChartType.LineStacked &&
                chartKind != OfficeChartType.LineStacked100 &&
                chartKind != OfficeChartType.Scatter &&
                chartKind != OfficeChartType.Radar) {
                return false;
            }

            bool sawSeries = false;
            bool sawHiddenMarker = false;
            foreach (OfficeXmlNode seriesElement in chartElement.Elements().Where(element => element.LocalName == "ser")) {
                sawSeries = true;
                if (!IsNativeWordChartSeriesMarkerHidden(seriesElement)) {
                    return false;
                }

                sawHiddenMarker = true;
            }

            return sawSeries && sawHiddenMarker;
        }

        // Stored symbols and sizes win; series without one take Office's automatic marker sequence.
        private static OfficeChartMarkerShape? GetNativeWordChartMarkerShape(OfficeXmlNode seriesElement, OfficeXmlNode chartElement, OfficeChartType chartKind) {
            if (!IsNativeWordLineLikeChart(chartKind) || IsNativeWordChartSeriesMarkerHidden(seriesElement)) return null;
            string? symbol = GetNativeWordChartMarkerSymbol(seriesElement);
            if (symbol != null && symbol != "auto" && symbol.Length > 0 && char.IsLetter(symbol[0]) &&
                Enum.TryParse(symbol, ignoreCase: true, out OfficeChartMarkerShape stored))
                return stored;
            return HasNativeWordChartGroupMarkers(chartElement)
                ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ResolveMarker(
                    GetNativeWordChartIndex(seriesElement) ?? 0U)
                : null;
        }

        /// <summary>The <c>c:symbol</c> token of a series' <c>c:marker</c>, or null.</summary>
        private static string? GetNativeWordChartMarkerSymbol(OfficeXmlNode seriesElement) =>
            GetNativeWordChartChild(seriesElement.Child(C, "marker"), "symbol").Attr("val");

        private static OfficeLength? GetNativeWordChartMarkerSize(OfficeXmlNode seriesElement, OfficeXmlNode chartElement, OfficeChartType chartKind) {
            string? sizeText = GetNativeWordChartChild(seriesElement.Child(C, "marker"), "size").Attr("val");
            byte? size = sizeText != null && byte.TryParse(sizeText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out byte parsedSize) ? parsedSize : null;
            if (size.HasValue) return OfficeLength.Points(size.Value);
            string? symbol = GetNativeWordChartMarkerSymbol(seriesElement);
            return IsNativeWordLineLikeChart(chartKind) && (symbol == null || symbol == "auto") &&
                HasNativeWordChartGroupMarkers(chartElement)
                ? OfficeLength.Points(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.AutomaticMarkerSize)
                : (OfficeLength?)null;
        }

        // Office draws automatic markers on line series without a stored marker even when the
        // group's c:marker is 0.
        private static bool HasNativeWordChartGroupMarkers(OfficeXmlNode chartElement) {
            if (chartElement.Is(C, "lineChart")) return true;
            if (chartElement.Is(C, "scatterChart")) {
                string? style = chartElement.Child(C, "scatterStyle").Attr("val");
                return style != "line" && style != "smooth";
            }
            return chartElement.Is(C, "radarChart") && chartElement.Child(C, "radarStyle").Attr("val") == "marker";
        }

        private static bool IsNativeWordChartSeriesMarkerHidden(OfficeXmlNode seriesElement) =>
            GetNativeWordChartMarkerSymbol(seriesElement) == "none";

        private static string? GetNativeWordChartDataLabelNumberFormat(OfficeXmlNode labels) {
            string? formatCode = GetNativeWordChartChild(labels, "numFmt").Attr("formatCode");
            return string.IsNullOrWhiteSpace(formatCode) ? null : formatCode;
        }

        private static OfficeChartDataLabelPosition GetNativeWordChartDataLabelPosition(OfficeXmlNode labels) =>
            GetNativeWordChartChild(labels, "dLblPos").Attr("val") switch {
                "ctr" => OfficeChartDataLabelPosition.Center,
                "inBase" => OfficeChartDataLabelPosition.InsideBase,
                "inEnd" => OfficeChartDataLabelPosition.InsideEnd,
                "outEnd" => OfficeChartDataLabelPosition.OutsideEnd,
                "l" => OfficeChartDataLabelPosition.Left,
                "r" => OfficeChartDataLabelPosition.Right,
                "t" => OfficeChartDataLabelPosition.Top,
                "b" => OfficeChartDataLabelPosition.Bottom,
                _ => OfficeChartDataLabelPosition.BestFit
            };

        private static bool HasNativeWordChartLegend(OfficeXmlNode chart) =>
            chart.Child(C, "legend").Exists;

        private static OfficeChartLegendPosition GetNativeOfficeChartLegendPosition(OfficeXmlNode chart) =>
            GetNativeWordChartChild(chart.Child(C, "legend"), "legendPos").Attr("val") switch {
                "l" => OfficeChartLegendPosition.Left,
                "t" => OfficeChartLegendPosition.Top,
                "b" => OfficeChartLegendPosition.Bottom,
                _ => OfficeChartLegendPosition.Right
            };

        private static int? GetNativeWordChartMaximumCategoryAxisLabels(OfficeXmlNode chartElement, OfficeXmlNode plotArea, int categoryCount) {
            if (categoryCount <= 0) {
                return null;
            }

            uint? skip = GetNativeWordChartCategoryAxisTickLabelSkip(chartElement, plotArea);
            if (!skip.HasValue || skip.Value <= 1U) {
                return null;
            }

            return Math.Max(1, (int)Math.Ceiling(categoryCount / (double)skip.Value));
        }

        private static uint? GetNativeWordChartCategoryAxisTickLabelSkip(OfficeXmlNode chartElement, OfficeXmlNode plotArea) {
            var chartAxisIds = new HashSet<uint>(GetNativeWordChartAxisIds(chartElement));

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartCategoryAxis)) {
                uint? axisId = GetNativeWordChartAxisId(axis);
                if (axisId.HasValue && chartAxisIds.Contains(axisId.Value)) {
                    return GetNativeWordChartTickLabelSkip(axis);
                }
            }

            foreach (OfficeXmlNode axis in plotArea.Elements(C, NativeWordChartCategoryAxis)) {
                uint? skip = GetNativeWordChartTickLabelSkip(axis);
                if (skip.HasValue) {
                    return skip;
                }
            }

            return null;
        }

        private static uint? GetNativeWordChartTickLabelSkip(OfficeXmlNode axis) {
            foreach (OfficeXmlNode skip in axis.DescendantsSkippingExtLst(C, "tickLblSkip")) {
                int? value = ParseInt32(skip.Attr("val"));
                if (value.HasValue && value.Value > 1) {
                    return (uint)value.Value;
                }
            }

            return null;
        }

        private static bool IsNativeWordHorizontalBarChart(OfficeChartType chartKind) =>
            chartKind == OfficeChartType.BarClustered ||
            chartKind == OfficeChartType.BarStacked ||
            chartKind == OfficeChartType.BarStacked100;

        private sealed class NativeWordChartDataLabelState {
            public OfficeXmlNode Labels { get; set; }
            public IReadOnlyCollection<int>? SeriesIndexes { get; set; }
            public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? PointIndexes { get; set; }
            public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? HiddenPointIndexes { get; set; }
            public bool ShowValue { get; set; }
            public bool ShowPercent { get; set; }
            public bool ShowCategoryName { get; set; }
            public bool ShowSeriesName { get; set; }
        }

        private static NativeWordChartDataLabelState GetNativeWordChartDataLabels(OfficeXmlNode chartElement, OfficeChartType chartKind) {
            var state = new NativeWordChartDataLabelState();
            List<int> renderableSeriesIndexes = GetNativeWordRenderableSeriesIndexes(chartElement, chartKind);
            OfficeXmlNode chartLabels = chartElement.Child(C, "dLbls");
            if (chartLabels.Exists) {
                state.Labels = chartLabels;
                state.ShowValue = HasNativeWordChartDataLabelFlag(NativeWordChartShowValue, chartLabels);
                state.ShowPercent = HasNativeWordChartDataLabelFlag(NativeWordChartShowPercent, chartLabels);
                state.ShowCategoryName = HasNativeWordChartDataLabelFlag(NativeWordChartShowCategoryName, chartLabels);
                state.ShowSeriesName = HasNativeWordChartDataLabelFlag(NativeWordChartShowSeriesName, chartLabels);
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
            foreach ((OfficeXmlNode seriesElement, _) in GetNativeWordOrderedSeriesElements(chartElement)) {
                if (!IsNativeWordChartSeriesRenderable(seriesElement, chartKind)) {
                    continue;
                }

                OfficeXmlNode labels = seriesElement.Child(C, "dLbls");
                if (labels.Exists) {
                    if (HasNativeWordChartDataLabelFlags(labels)) {
                        state.Labels = labels;
                        state.ShowValue |= HasNativeWordChartDataLabelFlag(NativeWordChartShowValue, labels);
                        state.ShowPercent |= HasNativeWordChartDataLabelFlag(NativeWordChartShowPercent, labels);
                        state.ShowCategoryName |= HasNativeWordChartDataLabelFlag(NativeWordChartShowCategoryName, labels);
                        state.ShowSeriesName |= HasNativeWordChartDataLabelFlag(NativeWordChartShowSeriesName, labels);
                        seriesIndexes.Add(seriesIndex);
                    } else if (state.Labels.IsNone) {
                        state.Labels = labels;
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

        private static List<int> GetNativeWordRenderableSeriesIndexes(OfficeXmlNode chartElement, OfficeChartType chartKind) {
            var indexes = new List<int>();
            int seriesIndex = 0;
            foreach (OfficeXmlNode seriesElement in chartElement.Elements().Where(element => element.LocalName == "ser")) {
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

        private static Dictionary<int, IReadOnlyCollection<int>>? GetNativeWordChartDataLabelPointIndexes(OfficeXmlNode labels, IReadOnlyList<int> seriesIndexes, bool deleted) {
            var indexes = new Dictionary<int, IReadOnlyCollection<int>>();
            foreach (int seriesIndex in seriesIndexes) {
                AddNativeWordChartDataLabelPointIndexes(labels, seriesIndex, indexes, deleted);
            }

            return indexes.Count == 0 ? null : indexes;
        }

        private static void AddNativeWordChartDataLabelPointIndexes(OfficeXmlNode labels, int seriesIndex, Dictionary<int, IReadOnlyCollection<int>> indexes, bool deleted) {
            var pointIndexes = new List<int>();
            foreach (OfficeXmlNode label in labels.Children(C, "dLbl")) {
                uint? index = GetNativeWordChartIndex(label);
                if (!index.HasValue) {
                    continue;
                }

                EnsureNativeWordChartPointIndex(index.Value);
                bool isDeleted = IsNativeWordChartBooleanOn(label.Child(C, "delete"));
                if (deleted == isDeleted && (deleted || HasNativeWordChartDataLabelFlags(label))) {
                    pointIndexes.Add((int)index.Value);
                }
            }

            if (pointIndexes.Count > 0) {
                indexes[seriesIndex] = pointIndexes;
            }
        }

        private const string NativeWordChartShowValue = "showVal";
        private const string NativeWordChartShowPercent = "showPercent";
        private const string NativeWordChartShowCategoryName = "showCatName";
        private const string NativeWordChartShowSeriesName = "showSerName";

        private static bool HasNativeWordChartDataLabelFlags(OfficeXmlNode labels) =>
            HasNativeWordChartDataLabelFlag(NativeWordChartShowValue, labels) ||
            HasNativeWordChartDataLabelFlag(NativeWordChartShowPercent, labels) ||
            HasNativeWordChartDataLabelFlag(NativeWordChartShowCategoryName, labels) ||
            HasNativeWordChartDataLabelFlag(NativeWordChartShowSeriesName, labels);

        private static bool HasNativeWordChartDataLabelFlag(string flagName, OfficeXmlNode labels) {
            if (labels.IsNone) {
                return false;
            }

            OfficeXmlNode value = labels.Child(C, flagName);
            if (IsNativeWordChartBooleanOn(value)) {
                return true;
            }

            return labels.Elements(C, "dLbl").Any(label => IsNativeWordChartBooleanOn(label.Child(C, flagName)));
        }

        private static bool IsNativeWordChartLegendOverlay(OfficeXmlNode chart) =>
            IsNativeWordChartBooleanOn(GetNativeWordChartChild(chart.Child(C, "legend"), "overlay"));

        private static bool IsNativeWordChartTitleOverlay(OfficeXmlNode chart) =>
            IsNativeWordChartBooleanOn(GetNativeWordChartChild(chart.Child(C, "title"), "overlay"));

        // A c:BooleanType element: on when present without c:val or with a true c:val.
        private static bool IsNativeWordChartBooleanOn(OfficeXmlNode value) =>
            value.Exists && (value.Attr("val") is not string text || ParseNativeWordChartBoolean(text) == true);

        /// <summary>An <c>xsd:boolean</c> value as the SDK's <c>BooleanValue</c> reads it; a value it can't parse throws as the SDK does.</summary>
        private static bool ParseNativeWordChartBoolean(string text) =>
            ParseBoolean(text) ?? throw new FormatException("The text value is not a valid boolean.");

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
