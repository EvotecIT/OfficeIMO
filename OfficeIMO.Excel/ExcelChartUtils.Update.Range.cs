using System.Globalization;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private const int MaxChartDataPoints = 1_000_000;

        private sealed class ChartDataPointBudget {
            private long _remaining;
            private long _remainingSourceReads;

            internal ChartDataPointBudget() : this(MaxChartDataPoints) { }

            internal ChartDataPointBudget(long maximum) {
                _remaining = maximum;
                _remainingSourceReads = maximum;
            }

            internal bool CanCharge(long count) => count >= 0 && count <= _remaining;

            internal bool TryCharge(long count) {
                if (!CanCharge(count)) {
                    return false;
                }
                _remaining -= count;
                return true;
            }

            internal bool TryChargeSourceRead(long count) {
                if (count < 0 || count > _remainingSourceReads) return false;
                _remainingSourceReads -= count;
                return true;
            }

            internal void RestoreUnusedSourceReads(long reserved, long consumed) {
                long unused = reserved - consumed;
                if (unused > 0) {
                    _remainingSourceReads += unused;
                }
            }
        }

        // The plot area of a chart part, or none.
        private static OfficeXmlNode GetPlotArea(ExcelPartRef chartPart) => chartPart.Root.Child(C, "chart").Child(C, "plotArea");

        // Whether a c:ser is a scatter series, which holds X and Y values instead of categories and values.
        private static bool IsScatterSeries(OfficeXmlNode series) => series.Parent.Is(C, "scatterChart");

        // Whether a c:ser is a scatter or bubble series.
        private static bool IsXySeries(OfficeXmlNode series) => series.Parent.Is(C, "scatterChart") || series.Parent.Is(C, "bubbleChart");

        // The c:f text of a c:numRef or c:strRef under a data source element, or null.
        private static string? ReferenceFormula(OfficeXmlNode source, string referenceName) =>
            source.Child(C, referenceName).ChildText(C, "f");

        internal static ExcelChartDataRange? TryExtractDataRange(ExcelPartRef chartPart, int maximumSeries = int.MaxValue) {
            OfficeXmlNode plotArea = GetPlotArea(chartPart);
            if (plotArea.IsNone) return null;

            IReadOnlyList<OfficeXmlNode> seriesList = GetChartSeries(plotArea, maximumSeries);

            if (seriesList.Count == 0) return null;

            var series = seriesList[0];
            string? catFormula;
            string? valFormula;
            if (IsScatterSeries(series)) {
                catFormula = ReferenceFormula(series.Child(C, "xVal"), "numRef");
                valFormula = ReferenceFormula(series.Child(C, "yVal"), "numRef");
            } else {
                OfficeXmlNode categoryAxisData = series.Child(C, "cat");
                catFormula = ReferenceFormula(categoryAxisData, "strRef")
                    ?? ReferenceFormula(categoryAxisData, "numRef");
                valFormula = ReferenceFormula(series.Child(C, "val"), "numRef");
            }

            if (!TryParseSheetQualifiedRange(catFormula, out var sheetName, out var catRange)) return null;
            if (!TryParseSheetQualifiedRange(valFormula, out var sheetNameValues, out var valRange)) return null;

            if (!string.Equals(sheetName, sheetNameValues, StringComparison.OrdinalIgnoreCase)) return null;
            if (!TryParseRange(catRange, out int r1, out int c1, out int r2, out int c2)) return null;
            if (!TryParseRange(valRange, out int valR1, out int valC1, out int valR2, out int valC2)) return null;

            int categoryRows = r2 - r1 + 1;
            int categoryColumns = c2 - c1 + 1;
            int valueRows = valR2 - valR1 + 1;
            int valueColumns = valC2 - valC1 + 1;
            bool horizontal = categoryRows == 1 && valueRows == 1 && valueColumns == categoryColumns &&
                valC1 == c1 && valC2 == c2 && valR1 == r1 + 1;
            int categoryCount = horizontal ? categoryColumns : categoryRows;
            if (categoryCount <= 0 ||
                (long)categoryCount + ((long)categoryCount * seriesList.Count) > MaxChartDataPoints) return null;

            if (horizontal) {
                if (valR1 != r1 + 1 || !HorizontalSeriesRangesAreContiguous(seriesList, sheetName, valR1, c1, c2)) {
                    return null;
                }

                bool hasHeaderColumn = HasHorizontalSeriesNameColumn(seriesList, sheetName, c1 - 1, valR1);
                int horizontalStartColumn = hasHeaderColumn ? c1 - 1 : c1;
                return new ExcelChartDataRange(
                    sheetName,
                    r1,
                    horizontalStartColumn,
                    categoryCount,
                    seriesList.Count,
                    hasHeaderColumn,
                    ExcelChartDataOrientation.Horizontal);
            }

            if (!VerticalSeriesRangesAreContiguous(seriesList, sheetName, r1, r2, c1 + 1)) {
                return null;
            }

            int headerRow = r1 - 1;
            bool hasHeaderRow = HasVerticalSeriesNameRow(seriesList, sheetName, headerRow);
            int startRow = hasHeaderRow ? headerRow : r1;
            int startColumn = c1;

            return new ExcelChartDataRange(sheetName, startRow, startColumn, categoryCount, seriesList.Count, hasHeaderRow);
        }

        internal static bool HasCanonicalWorksheetReferences(ExcelPartRef chartPart, ExcelChartDataRange range) {
            OfficeXmlNode plot = GetPlotArea(chartPart);
            if (plot.IsNone) return false;
            IReadOnlyList<OfficeXmlNode> series = GetChartSeries(plot, range.SeriesCount);
            if (series.Count != range.SeriesCount) return false;
            for (int index = 0; index < series.Count; index++) {
                OfficeXmlNode item = series[index];
                string? categories = ReferenceFormula(item.Child(C, "cat"), "strRef")
                    ?? ReferenceFormula(item.Child(C, "cat"), "numRef");
                if (!MatchesWorksheetRange(categories, range.SheetName,
                    range.CategoryStartRow, range.CategoryStartColumn,
                    range.CategoryEndRow, range.CategoryEndColumn)) return false;
                OfficeXmlNode values = GetSeriesValuesReference(item);
                int valueFirstRow = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesStartRow : range.SeriesStartRow + index;
                int valueFirstColumn = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesStartColumn + index : range.SeriesStartColumn;
                int valueLastRow = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesEndRow : valueFirstRow;
                int valueLastColumn = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? valueFirstColumn : range.SeriesEndColumn;
                if (!MatchesWorksheetRange(values.ChildText(C, "f"), range.SheetName,
                    valueFirstRow, valueFirstColumn, valueLastRow, valueLastColumn)) return false;
                if (range.HasHeaderRow) {
                    int labelRow = range.Orientation == ExcelChartDataOrientation.Vertical
                        ? range.StartRow : valueFirstRow;
                    int labelColumn = range.Orientation == ExcelChartDataOrientation.Vertical
                        ? valueFirstColumn : range.StartColumn;
                    if (!TryParseSeriesNameCell(item, range.SheetName, out int actualRow, out int actualColumn)
                        || actualRow != labelRow || actualColumn != labelColumn) return false;
                }
            }
            return true;
        }

        private static bool MatchesWorksheetRange(string? formula, string sheetName,
            int firstRow, int firstColumn, int lastRow, int lastColumn) =>
            TryParseSheetQualifiedRange(formula, out string actualSheet, out string actualRange)
            && string.Equals(sheetName, actualSheet, StringComparison.OrdinalIgnoreCase)
            && TryParseRange(actualRange, out int row1, out int column1, out int row2, out int column2)
            && row1 == firstRow && column1 == firstColumn && row2 == lastRow && column2 == lastColumn;

        private static bool HorizontalSeriesRangesAreContiguous(IReadOnlyList<OfficeXmlNode> seriesList, string sheetName, int firstSeriesRow, int firstCategoryColumn, int lastCategoryColumn) {
            for (int i = 0; i < seriesList.Count; i++) {
                OfficeXmlNode reference = GetSeriesValuesReference(seriesList[i]);
                if (!TryParseSheetQualifiedRange(reference.ChildText(C, "f"), out string valuesSheet, out string valuesRange)) {
                    return false;
                }

                if (!string.Equals(sheetName, valuesSheet, StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                if (!TryParseRange(valuesRange, out int row1, out int column1, out int row2, out int column2)) {
                    return false;
                }

                if (row1 != firstSeriesRow + i || row2 != row1 || column1 != firstCategoryColumn || column2 != lastCategoryColumn) {
                    return false;
                }
            }

            return true;
        }

        private static bool VerticalSeriesRangesAreContiguous(IReadOnlyList<OfficeXmlNode> seriesList, string sheetName, int firstCategoryRow, int lastCategoryRow, int firstSeriesColumn) {
            for (int i = 0; i < seriesList.Count; i++) {
                OfficeXmlNode reference = GetSeriesValuesReference(seriesList[i]);
                if (!TryParseSheetQualifiedRange(reference.ChildText(C, "f"), out string valuesSheet, out string valuesRange)) {
                    return false;
                }

                if (!string.Equals(sheetName, valuesSheet, StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                if (!TryParseRange(valuesRange, out int row1, out int column1, out int row2, out int column2)) {
                    return false;
                }

                int expectedColumn = firstSeriesColumn + i;
                if (row1 != firstCategoryRow || row2 != lastCategoryRow || column1 != expectedColumn || column2 != expectedColumn) {
                    return false;
                }
            }

            return true;
        }

        private static bool HasVerticalSeriesNameRow(IReadOnlyList<OfficeXmlNode> seriesList, string sheetName, int headerRow) {
            if (headerRow <= 0) {
                return false;
            }

            foreach (var seriesElement in seriesList) {
                if (!TryParseSeriesNameCell(seriesElement, sheetName, out int nameRow, out _)) {
                    continue;
                }

                if (nameRow == headerRow) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasHorizontalSeriesNameColumn(IReadOnlyList<OfficeXmlNode> seriesList, string sheetName, int headerColumn, int firstSeriesRow) {
            if (headerColumn <= 0) {
                return false;
            }

            foreach (var seriesElement in seriesList) {
                if (!TryParseSeriesNameCell(seriesElement, sheetName, out int nameRow, out int nameColumn)) {
                    continue;
                }

                if (nameColumn == headerColumn && nameRow >= firstSeriesRow && nameRow < firstSeriesRow + seriesList.Count) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseSeriesNameCell(OfficeXmlNode seriesElement, string sheetName, out int row, out int column) {
            row = 0;
            column = 0;
            string? nameFormula = ReferenceFormula(seriesElement.Child(C, "tx"), "strRef");
            if (!TryParseSheetQualifiedRange(nameFormula, out var nameSheet, out var nameRange)) {
                return false;
            }
            if (!string.Equals(nameSheet, sheetName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }
            if (!TryParseRange(nameRange, out int nameR1, out int nameC1, out int nameR2, out int nameC2)) {
                return false;
            }
            if (nameR1 != nameR2 || nameC1 != nameC2) {
                return false;
            }

            row = nameR1;
            column = nameC1;
            return true;
        }

        // The c:numRef of a series' values (Y values for a scatter series), or none.
        private static OfficeXmlNode GetSeriesValuesReference(OfficeXmlNode seriesElement) {
            if (IsScatterSeries(seriesElement)) {
                return seriesElement.Child(C, "yVal").Child(C, "numRef");
            }

            return seriesElement.Child(C, "val").Child(C, "numRef");
        }

        // The chart elements whose c:ser children are read as series; bubble charts are read on their own.
        private static bool HoldsReadableSeries(OfficeXmlNode chartElement) {
            if (chartElement.NamespaceId != C) return false;
            switch (chartElement.LocalName) {
                case "barChart":
                case "bar3DChart":
                case "lineChart":
                case "line3DChart":
                case "areaChart":
                case "area3DChart":
                case "pieChart":
                case "pie3DChart":
                case "ofPieChart":
                case "doughnutChart":
                case "radarChart":
                case "stockChart":
                case "surface3DChart":
                case "surfaceChart":
                case "scatterChart":
                    return true;
                default:
                    return false;
            }
        }

        private static IReadOnlyList<OfficeXmlNode> GetChartSeries(OfficeXmlNode plotArea, int maximumSeries = int.MaxValue) {
            var series = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode chartElement in plotArea.Children()) {
                if (!HoldsReadableSeries(chartElement)) continue;
                foreach (OfficeXmlNode item in chartElement.Children(C, "ser")) {
                    if (series.Count >= maximumSeries) return Array.Empty<OfficeXmlNode>();
                    series.Add(item);
                }
            }

            return series
                .OrderBy(item => item.Child(C, "idx").UInt32Val() ?? uint.MaxValue)
                .ToList();
        }

        /// <param name="sheet">The worksheet holding the range.</param>
        /// <param name="range">The chart's source range.</param>
        /// <param name="keepBlankCells">Whether an empty value cell becomes NaN (missing) instead of zero, as a plotted chart needs.</param>
        internal static ExcelChartData? TryReadChartData(ExcelSheet sheet, ExcelChartDataRange range, bool keepBlankCells = false) {
            try {
                if (range.CategoryCount <= 0 || range.SeriesCount <= 0 ||
                    (long)range.CategoryCount + ((long)range.CategoryCount * range.SeriesCount) > MaxChartDataPoints) return null;
                if (sheet.Document.TryReadDeferredTabularChartData(sheet, range, out ExcelChartData? deferred, keepBlankCells)) return deferred;
                var categories = new List<string>(range.CategoryCount);
                for (int i = 0; i < range.CategoryCount; i++) {
                    int row = range.Orientation == ExcelChartDataOrientation.Vertical ? range.CategoryStartRow + i : range.CategoryStartRow;
                    int column = range.Orientation == ExcelChartDataOrientation.Vertical ? range.CategoryStartColumn : range.CategoryStartColumn + i;
                    if (sheet.TryGetCellText(row, column, out var text)) {
                        categories.Add(text ?? string.Empty);
                    } else {
                        categories.Add(string.Empty);
                    }
                }

                var series = new List<ExcelChartSeries>(range.SeriesCount);
                for (int s = 0; s < range.SeriesCount; s++) {
                    string name = $"Series {s + 1}";
                    if (range.HasHeaderRow) {
                        int headerRow = range.Orientation == ExcelChartDataOrientation.Vertical ? range.StartRow : range.SeriesStartRow + s;
                        int headerColumn = range.Orientation == ExcelChartDataOrientation.Vertical ? range.SeriesStartColumn + s : range.StartColumn;
                        if (sheet.TryGetCellText(headerRow, headerColumn, out var header) && !string.IsNullOrWhiteSpace(header)) {
                            name = header;
                        }
                    }

                    var values = new List<double>(range.CategoryCount);
                    for (int i = 0; i < range.CategoryCount; i++) {
                        int row = range.Orientation == ExcelChartDataOrientation.Vertical ? range.CategoryStartRow + i : range.SeriesStartRow + s;
                        int column = range.Orientation == ExcelChartDataOrientation.Vertical ? range.SeriesStartColumn + s : range.CategoryStartColumn + i;
                        if (sheet.TryGetCellText(row, column, out var raw)
                            && OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) {
                            values.Add(val);
                        } else {
                            values.Add(keepBlankCells && string.IsNullOrWhiteSpace(raw) ? double.NaN : 0d);
                        }
                    }
                    series.Add(new ExcelChartSeries(name, values));
                }

                return new ExcelChartData(categories, series);
            } catch {
                return null;
            }
        }

        internal static ExcelChartData? TryReadChartData(ExcelPartRef chartPart, ExcelSheet contextSheet) =>
            TryReadPartChartData(chartPart, contextSheet, new ChartDataPointBudget());

        private static ExcelChartData? TryReadPartChartData(
            ExcelPartRef chartPart,
            ExcelSheet contextSheet,
            ChartDataPointBudget pointBudget) {
            try {
                OfficeXmlNode plotArea = GetPlotArea(chartPart);
                if (plotArea.IsNone) {
                    return null;
                }

                OfficeXmlNode bubble = plotArea.Child(C, "bubbleChart");
                if (bubble.Exists) {
                    return TryReadBubbleChartData(plotArea, bubble, contextSheet, pointBudget);
                }

                IReadOnlyList<OfficeXmlNode> seriesList = GetChartSeries(plotArea);
                if (seriesList.Count == 0) {
                    return null;
                }

                if (!TryReadCategoryValues(seriesList[0], contextSheet, pointBudget, out IReadOnlyList<string>? categories) || categories == null || categories.Count == 0) {
                    return null;
                }

                var series = new List<ExcelChartSeries>(seriesList.Count);
                for (int i = 0; i < seriesList.Count; i++) {
                    OfficeXmlNode seriesElement = seriesList[i];
                    OfficeXmlNode valuesReference = GetSeriesValuesReference(seriesElement);
                    string name = TryReadSeriesName(seriesElement, contextSheet, out string? resolvedName) && !string.IsNullOrWhiteSpace(resolvedName)
                        ? resolvedName!
                        : $"Series {i + 1}";
                    IReadOnlyList<double>? xValues = null;
                    if (IsScatterSeries(seriesElement)) {
                        if (i == 0) {
                            xValues = ParseNumericCategories(categories);
                        } else {
                            TryReadScatterXValues(seriesElement, contextSheet, pointBudget, out xValues);
                        }
                    }

                    TryReadReferencedPlotValues(contextSheet, valuesReference.ChildText(C, "f"), pointBudget, out IReadOnlyList<double>? referencedValues);
                    IReadOnlyList<double>? cachedValues = null;
                    if (referencedValues == null ||
                        xValues != null && referencedValues.Count != xValues.Count) {
                        TryReadCachedNumberValues(valuesReference, pointBudget, out cachedValues);
                    }

                    IReadOnlyList<double>? values = referencedValues;
                    if (xValues != null &&
                        referencedValues != null &&
                        referencedValues.Count != xValues.Count &&
                        cachedValues != null &&
                        cachedValues.Count == xValues.Count) {
                        values = cachedValues;
                    }

                    values ??= cachedValues;
                    if (values == null) {
                        return null;
                    }

                    if (xValues == null && values.Count != categories.Count) {
                        return null;
                    }

                    if (xValues != null && xValues.Count != values.Count) {
                        return null;
                    }

                    ExcelChartSeries chartSeries = new ExcelChartSeries(name, values);
                    series.Add(xValues == null ? chartSeries : chartSeries.WithXValues(xValues));
                }

                return new ExcelChartData(categories, series);
            } catch {
                return null;
            }
        }

        // A scatter series' X values: a literal first, then the referenced cells, then the reference's cache.
        private static void TryReadScatterXValues(OfficeXmlNode scatterSeries, ExcelSheet contextSheet, ChartDataPointBudget pointBudget,
            out IReadOnlyList<double>? xValues) {
            OfficeXmlNode xSource = scatterSeries.Child(C, "xVal");
            OfficeXmlNode xReference = xSource.Child(C, "numRef");
            OfficeXmlNode xLiteral = xSource.Child(C, "numLit");
            if (!TryReadNumberLiteralValues(xLiteral, pointBudget, out xValues) &&
                !TryReadReferencedNumberValues(contextSheet, xReference.ChildText(C, "f"), pointBudget, out xValues)) {
                TryReadCachedNumberValues(xReference, pointBudget, out xValues);
            }
        }

        private static bool TryReadCategoryValues(OfficeXmlNode seriesElement, ExcelSheet contextSheet, ChartDataPointBudget pointBudget, out IReadOnlyList<string>? categories) {
            categories = null;
            if (IsScatterSeries(seriesElement)) {
                TryReadScatterXValues(seriesElement, contextSheet, pointBudget, out IReadOnlyList<double>? numericValues);

                if (numericValues == null) {
                    return false;
                }

                categories = numericValues.Select(InvariantNumberText.Get).ToArray();
                return categories.Count > 0;
            }

            OfficeXmlNode categoryAxisData = seriesElement.Child(C, "cat");
            OfficeXmlNode stringReference = categoryAxisData.Child(C, "strRef");
            if (TryReadReferencedTextValues(contextSheet, stringReference.ChildText(C, "f"), pointBudget, out categories)) {
                return true;
            }

            if (TryReadCachedStringValues(stringReference, pointBudget, out categories)) {
                return true;
            }

            OfficeXmlNode numberReference = categoryAxisData.Child(C, "numRef");
            if (!TryReadReferencedNumberValues(contextSheet, numberReference.ChildText(C, "f"), pointBudget, out IReadOnlyList<double>? numberValues)) {
                TryReadCachedNumberValues(numberReference, pointBudget, out numberValues);
            }

            if (numberValues == null) {
                return false;
            }

            categories = numberValues.Select(InvariantNumberText.Get).ToArray();
            return categories.Count > 0;
        }

        private static bool TryReadSeriesName(OfficeXmlNode seriesElement, ExcelSheet contextSheet, out string? name) {
            name = null;
            var nameBudget = new ChartDataPointBudget(1L);
            OfficeXmlNode seriesText = seriesElement.Child(C, "tx");
            OfficeXmlNode reference = seriesText.Child(C, "strRef");
            if (TryReadReferencedTextValues(contextSheet, reference.ChildText(C, "f"), nameBudget, out IReadOnlyList<string>? referencedValues) && referencedValues != null && referencedValues.Count > 0) {
                name = referencedValues[0];
                return true;
            }

            nameBudget = new ChartDataPointBudget(1L);
            if (TryReadCachedStringValues(reference, nameBudget, out IReadOnlyList<string>? cachedValues) && cachedValues != null && cachedValues.Count > 0) {
                name = cachedValues[0];
                return true;
            }

            OfficeXmlNode literal = seriesText.Child(C, "strLit");
            nameBudget = new ChartDataPointBudget(1L);
            if (TryReadFirstStringLiteralValue(literal, nameBudget, out string? literalText) &&
                !string.IsNullOrWhiteSpace(literalText)) {
                name = literalText;
                return true;
            }

            string? directText = seriesText.ChildText(C, "v");
            if (!string.IsNullOrWhiteSpace(directText)) {
                name = directText;
                return true;
            }

            return false;
        }

        private static bool TryReadFirstStringLiteralValue(
            OfficeXmlNode literal,
            ChartDataPointBudget pointBudget,
            out string? value) {
            value = null;
            uint selectedIndex = uint.MaxValue;
            bool selected = false;
            if (literal.IsNone) {
                return false;
            }

            foreach (OfficeXmlNode point in literal.Children(C, "pt")) {
                if (!pointBudget.TryCharge(1)) {
                    value = null;
                    return false;
                }

                uint index = point.UInt32Attribute("idx") ?? uint.MaxValue;
                if (!selected || index < selectedIndex) {
                    selected = true;
                    selectedIndex = index;
                    value = point.ChildText(C, "v");
                }
            }

            return selected;
        }

        internal static ExcelChartData ApplyChartSeriesTypes(ExcelPartRef chartPart, ExcelChartData data, OfficeChartType defaultType) {
            OfficeXmlNode plotArea = GetPlotArea(chartPart);
            if (plotArea.IsNone || data.Series.Count == 0) {
                return data;
            }

            var typedSeries = new List<(OfficeXmlNode Series, OfficeChartType ChartType, OfficeChartAxisGroup AxisGroup)>();
            foreach (OfficeXmlNode chartElement in plotArea.Children()) {
                if (!TryGetChartElementType(chartElement, out OfficeChartType chartType)) {
                    continue;
                }

                OfficeChartAxisGroup axisGroup = GetChartAxisGroup(plotArea, chartElement);
                foreach (OfficeXmlNode child in chartElement.Children()) {
                    if (child.Child(C, "idx").UInt32Val() != null) {
                        typedSeries.Add((child, chartType, axisGroup));
                    }
                }
            }

            var literalNames = GetChartSeries(plotArea, data.Series.Count)
                .Select((item, index) => new { index, Name = item.Child(C, "tx").ChildText(C, "v") })
                .Where(item => item.Name != null)
                .ToDictionary(item => item.index, item => item.Name!);
            if (typedSeries.Count == 0 && literalNames.Count == 0) {
                return data;
            }

            Dictionary<int, (OfficeChartType ChartType, OfficeChartAxisGroup AxisGroup)> seriesTypes = typedSeries
                .OrderBy(item => item.Series.Child(C, "idx").UInt32Val() ?? uint.MaxValue)
                .Select((item, index) => new { index, item.ChartType, item.AxisGroup })
                .ToDictionary(item => item.index, item => (item.ChartType, item.AxisGroup));

            var series = new List<ExcelChartSeries>(data.Series.Count);
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                OfficeChartType? seriesType = current.ChartType;
                OfficeChartAxisGroup axisGroup = current.AxisGroup;
                // Literal names also belong to native chart families whose type
                // projection is unsupported. Linked headers use current sheet data.
                string name = literalNames.TryGetValue(i, out string? literalName) ? literalName : current.Name;
                if (seriesTypes.TryGetValue(i, out var authoredSeries)) {
                    seriesType = authoredSeries.ChartType;
                    axisGroup = authoredSeries.AxisGroup;
                }

                series.Add(new ExcelChartSeries(name, current.Values, seriesType, axisGroup).WithXValues(current.XValues));
            }

            return new ExcelChartData(data.Categories, series);
        }

        private static OfficeChartAxisGroup GetChartAxisGroup(OfficeXmlNode plotArea, OfficeXmlNode chartElement) {
            var axisIds = new HashSet<uint>();
            foreach (OfficeXmlNode axis in chartElement.Children(C, "axId")) {
                if (axis.UInt32Val() is uint id) axisIds.Add(id);
            }
            foreach (OfficeXmlNode axis in plotArea.Children(C, "valAx")) {
                uint? id = axis.Child(C, "axId").UInt32Val();
                string? position = axis.ChildVal(C, "axPos");
                if (id != null && axisIds.Contains(id.Value) && (position == AxisPositionRight || position == AxisPositionTop)) {
                    return OfficeChartAxisGroup.Secondary;
                }
            }
            return OfficeChartAxisGroup.Primary;
        }

        internal static ExcelChartData ApplyScatterSeriesXValues(ExcelPartRef chartPart, ExcelChartData data, ExcelSheet contextSheet) {
            OfficeXmlNode plotArea = GetPlotArea(chartPart);
            if (plotArea.IsNone || data.Series.Count == 0) {
                return data;
            }

            var pointBudget = new ChartDataPointBudget();
            var scatterValuesByIndex = new Dictionary<int, (IReadOnlyList<double> XValues, IReadOnlyList<double> YValues)>();
            int seriesOrder = 0;
            foreach (OfficeXmlNode scatterSeries in GetChartSeries(plotArea)) {
                int index = seriesOrder++;
                if (!IsScatterSeries(scatterSeries)) {
                    continue;
                }

                if (index >= data.Series.Count) {
                    break;
                }

                TryReadScatterXValues(scatterSeries, contextSheet, pointBudget, out IReadOnlyList<double>? xValues);

                if (xValues == null) {
                    continue;
                }

                IReadOnlyList<double>? yValues = data.Series[index].Values;
                if (yValues.Count != xValues.Count) {
                    OfficeXmlNode yReference = scatterSeries.Child(C, "yVal").Child(C, "numRef");
                    TryReadReferencedPlotValues(contextSheet, yReference.ChildText(C, "f"), pointBudget, out IReadOnlyList<double>? referencedYValues);
                    IReadOnlyList<double>? cachedYValues = null;
                    if (referencedYValues == null || referencedYValues.Count != xValues.Count) {
                        TryReadCachedNumberValues(yReference, pointBudget, out cachedYValues);
                    }

                    yValues = referencedYValues?.Count == xValues.Count
                        ? referencedYValues
                        : cachedYValues?.Count == xValues.Count ? cachedYValues : null;
                }

                if (yValues != null && yValues.Count == xValues.Count) {
                    scatterValuesByIndex[index] = (xValues, yValues);
                }
            }

            if (scatterValuesByIndex.Count == 0) {
                return data;
            }

            var series = new List<ExcelChartSeries>(data.Series.Count);
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                if (!scatterValuesByIndex.TryGetValue(i, out var scatterValues)) {
                    series.Add(current);
                    continue;
                }

                series.Add(current.WithValuesAndXValues(scatterValues.YValues, scatterValues.XValues));
            }

            return new ExcelChartData(data.Categories, series);
        }

        private static bool TryReadCachedNumberValues(OfficeXmlNode reference, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            values = null;
            OfficeXmlNode cache = reference.Child(C, "numCache");
            return TryReadNumberPointNodes(cache, pointBudget, out values);
        }

        private static bool TryReadNumberLiteralValues(OfficeXmlNode literal, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            return TryReadNumberPointNodes(literal, pointBudget, out values);
        }

        private static bool TryReadNumberPointNodes(OfficeXmlNode container, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            values = null;
            if (container.IsNone) {
                return false;
            }

            uint? pointCount = container.Child(C, "ptCount").UInt32Val();
            int actualPointCount = CountChildren(container, "pt", MaxChartDataPoints + 1);
            if (actualPointCount > MaxChartDataPoints) {
                return false;
            }
            if (pointCount.HasValue) {
                long chargedPoints = Math.Max((long)pointCount.Value, actualPointCount);
                if (pointCount.Value > MaxChartDataPoints || !pointBudget.TryCharge(chargedPoints)) {
                    return false;
                }

                var indexedValues = Enumerable.Repeat(double.NaN, (int)pointCount.Value).ToArray();
                int nextUnindexedIndex = 0;
                bool hasValues = false;
                foreach (OfficeXmlNode point in container.Children(C, "pt")) {
                    string? text = point.ChildText(C, "v");
                    if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                        return false;
                    }

                    uint index = point.UInt32Attribute("idx") ?? (uint)nextUnindexedIndex++;
                    if (index >= pointCount.Value) {
                        continue;
                    }

                    indexedValues[(int)index] = value;
                    hasValues = true;
                }

                values = indexedValues;
                return hasValues;
            }

            if (!pointBudget.TryCharge(actualPointCount)) {
                return false;
            }
            var numericValues = new List<double>(actualPointCount);
            foreach (OfficeXmlNode point in container.Children(C, "pt").ToList().OrderBy(point => point.UInt32Attribute("idx") ?? uint.MaxValue)) {
                string? text = point.ChildText(C, "v");
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                    return false;
                }

                numericValues.Add(value);
            }

            values = numericValues;
            return numericValues.Count > 0;
        }

        // The number of children with this name, counted up to a limit.
        private static int CountChildren(OfficeXmlNode container, string localName, int limit) {
            int count = 0;
            foreach (OfficeXmlNode _ in container.Children(C, localName)) {
                if (++count >= limit) break;
            }
            return count;
        }

        private static bool TryReadCachedStringValues(OfficeXmlNode reference, ChartDataPointBudget pointBudget, out IReadOnlyList<string>? values) {
            values = null;
            OfficeXmlNode cache = reference.Child(C, "strCache");
            if (cache.IsNone) {
                return false;
            }

            int actualPointCount = CountChildren(cache, "pt", MaxChartDataPoints + 1);
            if (actualPointCount > MaxChartDataPoints || !pointBudget.TryCharge(actualPointCount)) return false;
            string[] cachedValues = cache.Children(C, "pt").ToList()
                .OrderBy(point => point.UInt32Attribute("idx") ?? uint.MaxValue)
                .Select(point => point.ChildText(C, "v") ?? string.Empty)
                .ToArray();
            values = cachedValues;
            return values.Count > 0;
        }

        private static bool TryReadReferencedTextValues(ExcelSheet contextSheet, string? formula, ChartDataPointBudget pointBudget, out IReadOnlyList<string>? values) {
            values = null;
            if (!TryParseSheetQualifiedRange(formula, out string sheetName, out string range) ||
                !TryParseRange(range, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            ExcelSheet sheet;
            try {
                sheet = contextSheet.Document[sheetName];
            } catch {
                return false;
            }

            long requestedPointCount = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (!pointBudget.CanCharge(requestedPointCount) ||
                !pointBudget.TryChargeSourceRead(requestedPointCount)) {
                return false;
            }

            var textValues = new List<string>();
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    textValues.Add(sheet.TryGetCellText(row, column, out string? text) ? text ?? string.Empty : string.Empty);
                }
            }

            if (textValues.Count == 0 || !pointBudget.TryCharge(requestedPointCount)) return false;
            values = textValues;
            return true;
        }

        private static bool TryReadReferencedNumberValues(ExcelSheet contextSheet, string? formula, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) =>
            TryReadReferencedNumbers(contextSheet, formula, pointBudget, 0D, out values);

        // Plotted values keep an empty cell as missing (NaN) so the chart's blank-cell rule decides how it is drawn.
        private static bool TryReadReferencedPlotValues(ExcelSheet contextSheet, string? formula, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) =>
            TryReadReferencedNumbers(contextSheet, formula, pointBudget, double.NaN, out values);

        private static bool TryReadReferencedNumbers(ExcelSheet contextSheet, string? formula, ChartDataPointBudget pointBudget,
            double blankValue, out IReadOnlyList<double>? values) {
            values = null;
            if (!TryParseSheetQualifiedRange(formula, out string sheetName, out string range) ||
                !TryParseRange(range, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            ExcelSheet sheet;
            try {
                sheet = contextSheet.Document[sheetName];
            } catch {
                return false;
            }

            long requestedPointCount = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (!pointBudget.CanCharge(requestedPointCount) ||
                !pointBudget.TryChargeSourceRead(requestedPointCount)) {
                return false;
            }

            var numericValues = new List<double>();
            long consumedSourceReads = 0;
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    consumedSourceReads++;
                    if (!sheet.TryGetCellText(row, column, out string? raw) ||
                        string.IsNullOrWhiteSpace(raw)) {
                        numericValues.Add(blankValue);
                        continue;
                    }

                    if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)) {
                        pointBudget.RestoreUnusedSourceReads(requestedPointCount, consumedSourceReads);
                        return false;
                    }

                    numericValues.Add(value);
                }
            }

            if (numericValues.Count == 0 || !pointBudget.TryCharge(requestedPointCount)) {
                return false;
            }
            values = numericValues;
            return true;
        }

        private static bool TryGetChartElementType(OfficeXmlNode element, out OfficeChartType chartType) {
            switch (element.NamespaceId == C ? element.LocalName : string.Empty) {
                case "barChart":
                    string direction = element.ChildVal(C, "barDir") ?? BarDirectionColumn;
                    string barGrouping = element.ChildVal(C, "grouping") ?? BarGroupingClustered;
                    if (direction == BarDirectionBar) {
                        chartType = barGrouping == GroupingPercentStacked
                            ? OfficeChartType.BarStacked100
                            : barGrouping == GroupingStacked ? OfficeChartType.BarStacked : OfficeChartType.BarClustered;
                    } else {
                        chartType = barGrouping == GroupingPercentStacked
                            ? OfficeChartType.ColumnStacked100
                            : barGrouping == GroupingStacked ? OfficeChartType.ColumnStacked : OfficeChartType.ColumnClustered;
                    }

                    return true;
                case "lineChart":
                    string lineGrouping = element.ChildVal(C, "grouping") ?? GroupingStandard;
                    chartType = lineGrouping == GroupingPercentStacked
                        ? OfficeChartType.LineStacked100
                        : lineGrouping == GroupingStacked ? OfficeChartType.LineStacked : OfficeChartType.Line;
                    return true;
                case "areaChart":
                    string areaGrouping = element.ChildVal(C, "grouping") ?? GroupingStandard;
                    chartType = areaGrouping == GroupingPercentStacked
                        ? OfficeChartType.AreaStacked100
                        : areaGrouping == GroupingStacked ? OfficeChartType.AreaStacked : OfficeChartType.Area;
                    return true;
                case "scatterChart":
                    chartType = OfficeChartType.Scatter;
                    return true;
                case "radarChart":
                    chartType = OfficeChartType.Radar;
                    return true;
                case "pieChart":
                    chartType = OfficeChartType.Pie;
                    return true;
                case "pie3DChart":
                    chartType = OfficeChartType.Pie3D;
                    return true;
                case "doughnutChart":
                    chartType = OfficeChartType.Doughnut;
                    return true;
                default:
                    chartType = default;
                    return false;
            }
        }
    }
}
