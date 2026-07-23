using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

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

        internal static ExcelChartDataRange? TryExtractDataRange(ChartPart chartPart) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) return null;

            IReadOnlyList<OpenXmlCompositeElement> seriesList = GetChartSeries(plotArea);

            if (seriesList.Count == 0) return null;

            var series = seriesList[0];
            string? catFormula;
            string? valFormula;
            if (series is ScatterChartSeries scatterSeries) {
                catFormula = scatterSeries.GetFirstChild<XValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
                valFormula = scatterSeries.GetFirstChild<YValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
            } else {
                CategoryAxisData? categoryAxisData = series.GetFirstChild<CategoryAxisData>();
                catFormula = categoryAxisData?
                    .GetFirstChild<StringReference>()?
                    .Formula?.Text
                    ?? categoryAxisData?
                        .GetFirstChild<NumberReference>()?
                        .Formula?.Text;
                valFormula = series.GetFirstChild<Values>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
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
            bool horizontal = categoryRows == 1 && categoryColumns > 1 && valueRows == 1 && valueColumns == categoryColumns && valC1 == c1 && valC2 == c2;
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

        private static bool HorizontalSeriesRangesAreContiguous(IReadOnlyList<OpenXmlCompositeElement> seriesList, string sheetName, int firstSeriesRow, int firstCategoryColumn, int lastCategoryColumn) {
            for (int i = 0; i < seriesList.Count; i++) {
                NumberReference? reference = GetSeriesValuesReference(seriesList[i]);
                if (!TryParseSheetQualifiedRange(reference?.Formula?.Text, out string valuesSheet, out string valuesRange)) {
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

        private static bool VerticalSeriesRangesAreContiguous(IReadOnlyList<OpenXmlCompositeElement> seriesList, string sheetName, int firstCategoryRow, int lastCategoryRow, int firstSeriesColumn) {
            for (int i = 0; i < seriesList.Count; i++) {
                NumberReference? reference = GetSeriesValuesReference(seriesList[i]);
                if (!TryParseSheetQualifiedRange(reference?.Formula?.Text, out string valuesSheet, out string valuesRange)) {
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

        private static bool HasVerticalSeriesNameRow(IReadOnlyList<OpenXmlCompositeElement> seriesList, string sheetName, int headerRow) {
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

        private static bool HasHorizontalSeriesNameColumn(IReadOnlyList<OpenXmlCompositeElement> seriesList, string sheetName, int headerColumn, int firstSeriesRow) {
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

        private static bool TryParseSeriesNameCell(OpenXmlCompositeElement seriesElement, string sheetName, out int row, out int column) {
            row = 0;
            column = 0;
            string? nameFormula = seriesElement.GetFirstChild<SeriesText>()?
                .GetFirstChild<StringReference>()?
                .Formula?.Text;
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

        private static NumberReference? GetSeriesValuesReference(OpenXmlCompositeElement seriesElement) {
            if (seriesElement is ScatterChartSeries scatterSeries) {
                return scatterSeries.GetFirstChild<YValues>()?.GetFirstChild<NumberReference>();
            }

            return seriesElement.GetFirstChild<Values>()?.GetFirstChild<NumberReference>();
        }

        private static IReadOnlyList<OpenXmlCompositeElement> GetChartSeries(PlotArea plotArea) {
            var series = new List<OpenXmlCompositeElement>();
            foreach (OpenXmlElement chartElement in plotArea.ChildElements) {
                switch (chartElement) {
                    case BarChart bar:
                        series.AddRange(bar.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Bar3DChart bar3D:
                        series.AddRange(bar3D.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case LineChart line:
                        series.AddRange(line.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Line3DChart line3D:
                        series.AddRange(line3D.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case AreaChart area:
                        series.AddRange(area.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Area3DChart area3D:
                        series.AddRange(area3D.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case PieChart pie:
                        series.AddRange(pie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Pie3DChart pie3D:
                        series.AddRange(pie3D.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case OfPieChart ofPie:
                        series.AddRange(ofPie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case DoughnutChart doughnut:
                        series.AddRange(doughnut.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case RadarChart radar:
                        series.AddRange(radar.Elements<RadarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case StockChart stock:
                        series.AddRange(stock.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Surface3DChart surface3D:
                        series.AddRange(surface3D.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case SurfaceChart surface:
                        series.AddRange(surface.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case ScatterChart scatter:
                        series.AddRange(scatter.Elements<ScatterChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                }
            }

            return series
                .OrderBy(item => item.GetFirstChild<ChartIndex>()?.Val?.Value ?? uint.MaxValue)
                .ToList();
        }

        internal static ExcelChartData? TryReadChartData(ExcelSheet sheet, ExcelChartDataRange range) {
            try {
                if (range.CategoryCount <= 0 || range.SeriesCount <= 0 ||
                    (long)range.CategoryCount + ((long)range.CategoryCount * range.SeriesCount) > MaxChartDataPoints) return null;
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
                            && double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) {
                            values.Add(val);
                        } else {
                            values.Add(0d);
                        }
                    }
                    series.Add(new ExcelChartSeries(name, values));
                }

                return new ExcelChartData(categories, series);
            } catch {
                return null;
            }
        }

        internal static ExcelChartData? TryReadChartData(ChartPart chartPart, ExcelSheet contextSheet) =>
            TryReadChartDataCore(chartPart, contextSheet, new ChartDataPointBudget());

        private static ExcelChartData? TryReadChartDataCore(
            ChartPart chartPart,
            ExcelSheet contextSheet,
            ChartDataPointBudget pointBudget) {
            try {
                var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
                var plotArea = chart?.GetFirstChild<PlotArea>();
                if (plotArea == null) {
                    return null;
                }

                IReadOnlyList<OpenXmlCompositeElement> seriesList = GetChartSeries(plotArea);
                if (seriesList.Count == 0) {
                    return null;
                }

                if (!TryReadCategoryValues(seriesList[0], contextSheet, pointBudget, out IReadOnlyList<string>? categories) || categories == null || categories.Count == 0) {
                    return null;
                }

                var series = new List<ExcelChartSeries>(seriesList.Count);
                for (int i = 0; i < seriesList.Count; i++) {
                    OpenXmlCompositeElement seriesElement = seriesList[i];
                    NumberReference? valuesReference = GetSeriesValuesReference(seriesElement);
                    string name = TryReadSeriesName(seriesElement, contextSheet, out string? resolvedName) && !string.IsNullOrWhiteSpace(resolvedName)
                        ? resolvedName!
                        : $"Series {i + 1}";
                    IReadOnlyList<double>? xValues = null;
                    if (seriesElement is ScatterChartSeries scatterSeries) {
                        if (i == 0) {
                            xValues = ParseNumericCategories(categories);
                        } else {
                            NumberReference? xReference = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberReference>();
                            NumberLiteral? xLiteral = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberLiteral>();
                            if (!TryReadNumberLiteralValues(xLiteral, pointBudget, out xValues) &&
                                !TryReadReferencedNumberValues(contextSheet, xReference?.Formula?.Text, pointBudget, out xValues)) {
                                TryReadCachedNumberValues(xReference, pointBudget, out xValues);
                            }
                        }
                    }

                    TryReadReferencedNumberValues(contextSheet, valuesReference?.Formula?.Text, pointBudget, out IReadOnlyList<double>? referencedValues);
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

        private static bool TryReadCategoryValues(OpenXmlCompositeElement seriesElement, ExcelSheet contextSheet, ChartDataPointBudget pointBudget, out IReadOnlyList<string>? categories) {
            categories = null;
            if (seriesElement is ScatterChartSeries scatterSeries) {
                NumberReference? xReference = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberReference>();
                NumberLiteral? xLiteral = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberLiteral>();
                if (!TryReadNumberLiteralValues(xLiteral, pointBudget, out IReadOnlyList<double>? numericValues) &&
                    !TryReadReferencedNumberValues(contextSheet, xReference?.Formula?.Text, pointBudget, out numericValues)) {
                    TryReadCachedNumberValues(xReference, pointBudget, out numericValues);
                }

                if (numericValues == null) {
                    return false;
                }

                categories = numericValues.Select(InvariantNumberText.Get).ToArray();
                return categories.Count > 0;
            }

            CategoryAxisData? categoryAxisData = seriesElement.GetFirstChild<CategoryAxisData>();
            StringReference? stringReference = categoryAxisData?.GetFirstChild<StringReference>();
            if (TryReadReferencedTextValues(contextSheet, stringReference?.Formula?.Text, pointBudget, out categories)) {
                return true;
            }

            if (TryReadCachedStringValues(stringReference, pointBudget, out categories)) {
                return true;
            }

            NumberReference? numberReference = categoryAxisData?.GetFirstChild<NumberReference>();
            if (!TryReadReferencedNumberValues(contextSheet, numberReference?.Formula?.Text, pointBudget, out IReadOnlyList<double>? numberValues)) {
                TryReadCachedNumberValues(numberReference, pointBudget, out numberValues);
            }

            if (numberValues == null) {
                return false;
            }

            categories = numberValues.Select(InvariantNumberText.Get).ToArray();
            return categories.Count > 0;
        }

        private static bool TryReadSeriesName(OpenXmlCompositeElement seriesElement, ExcelSheet contextSheet, out string? name) {
            name = null;
            var nameBudget = new ChartDataPointBudget(1L);
            SeriesText? seriesText = seriesElement.GetFirstChild<SeriesText>();
            StringReference? reference = seriesText?.GetFirstChild<StringReference>();
            if (TryReadReferencedTextValues(contextSheet, reference?.Formula?.Text, nameBudget, out IReadOnlyList<string>? referencedValues) && referencedValues != null && referencedValues.Count > 0) {
                name = referencedValues[0];
                return true;
            }

            nameBudget = new ChartDataPointBudget(1L);
            if (TryReadCachedStringValues(reference, nameBudget, out IReadOnlyList<string>? cachedValues) && cachedValues != null && cachedValues.Count > 0) {
                name = cachedValues[0];
                return true;
            }

            StringLiteral? literal = seriesText?.GetFirstChild<StringLiteral>();
            nameBudget = new ChartDataPointBudget(1L);
            if (TryReadFirstStringLiteralValue(literal, nameBudget, out string? literalText) &&
                !string.IsNullOrWhiteSpace(literalText)) {
                name = literalText;
                return true;
            }

            string? directText = seriesText?.GetFirstChild<NumericValue>()?.Text;
            if (!string.IsNullOrWhiteSpace(directText)) {
                name = directText;
                return true;
            }

            return false;
        }

        private static bool TryReadFirstStringLiteralValue(
            StringLiteral? literal,
            ChartDataPointBudget pointBudget,
            out string? value) {
            value = null;
            uint selectedIndex = uint.MaxValue;
            bool selected = false;
            if (literal == null) {
                return false;
            }

            foreach (StringPoint point in literal.Elements<StringPoint>()) {
                if (!pointBudget.TryCharge(1)) {
                    value = null;
                    return false;
                }

                uint index = point.Index?.Value ?? uint.MaxValue;
                if (!selected || index < selectedIndex) {
                    selected = true;
                    selectedIndex = index;
                    value = point.NumericValue?.Text;
                }
            }

            return selected;
        }

        internal static ExcelChartData ApplyChartSeriesTypes(ChartPart chartPart, ExcelChartData data, ExcelChartType defaultType) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null || data.Series.Count == 0) {
                return data;
            }

            var typedSeries = new List<(OpenXmlCompositeElement Series, ExcelChartType ChartType, ExcelChartAxisGroup AxisGroup)>();
            foreach (OpenXmlElement chartElement in plotArea.ChildElements) {
                if (!TryGetChartElementType(chartElement, out ExcelChartType chartType)) {
                    continue;
                }

                ExcelChartAxisGroup axisGroup = GetChartAxisGroup(plotArea, chartElement);
                foreach (OpenXmlElement child in chartElement.ChildElements) {
                    if (child is OpenXmlCompositeElement seriesElement && seriesElement.GetFirstChild<ChartIndex>()?.Val?.Value != null) {
                        typedSeries.Add((seriesElement, chartType, axisGroup));
                    }
                }
            }

            if (typedSeries.Count == 0) {
                return data;
            }

            Dictionary<int, (ExcelChartType ChartType, ExcelChartAxisGroup AxisGroup)> seriesTypes = typedSeries
                .OrderBy(item => item.Series.GetFirstChild<ChartIndex>()?.Val?.Value ?? uint.MaxValue)
                .Select((item, index) => new { index, item.ChartType, item.AxisGroup })
                .ToDictionary(item => item.index, item => (item.ChartType, item.AxisGroup));

            var series = new List<ExcelChartSeries>(data.Series.Count);
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                ExcelChartType? seriesType = current.ChartType;
                ExcelChartAxisGroup axisGroup = current.AxisGroup;
                if (seriesTypes.TryGetValue(i, out var authoredSeries)) {
                    seriesType = authoredSeries.ChartType;
                    axisGroup = authoredSeries.AxisGroup;
                }

                series.Add(new ExcelChartSeries(current.Name, current.Values, seriesType, axisGroup).WithXValues(current.XValues));
            }

            return new ExcelChartData(data.Categories, series);
        }

        private static ExcelChartAxisGroup GetChartAxisGroup(PlotArea plotArea, OpenXmlElement chartElement) {
            HashSet<uint> axisIds = new(chartElement.Elements<AxisId>()
                .Where(axis => axis.Val?.Value != null)
                .Select(axis => axis.Val!.Value));
            return plotArea.Elements<ValueAxis>().Any(axis =>
                       axis.AxisId?.Val?.Value != null && axisIds.Contains(axis.AxisId.Val.Value) &&
                       (axis.AxisPosition?.Val?.Value == AxisPositionValues.Right ||
                        axis.AxisPosition?.Val?.Value == AxisPositionValues.Top))
                ? ExcelChartAxisGroup.Secondary
                : ExcelChartAxisGroup.Primary;
        }

        internal static ExcelChartData ApplyScatterSeriesXValues(ChartPart chartPart, ExcelChartData data, ExcelSheet contextSheet) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null || data.Series.Count == 0) {
                return data;
            }

            var pointBudget = new ChartDataPointBudget();
            var scatterXValuesByIndex = new Dictionary<int, IReadOnlyList<double>>();
            int seriesOrder = 0;
            foreach (OpenXmlElement seriesElement in GetChartSeries(plotArea)) {
                int index = seriesOrder++;
                if (seriesElement is not ScatterChartSeries scatterSeries) {
                    continue;
                }

                if (index >= data.Series.Count) {
                    break;
                }

                NumberReference? xReference = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberReference>();
                NumberLiteral? xLiteral = scatterSeries.GetFirstChild<XValues>()?.GetFirstChild<NumberLiteral>();
                if (!TryReadNumberLiteralValues(xLiteral, pointBudget, out IReadOnlyList<double>? xValues) &&
                    !TryReadReferencedNumberValues(contextSheet, xReference?.Formula?.Text, pointBudget, out xValues)) {
                    TryReadCachedNumberValues(xReference, pointBudget, out xValues);
                }

                if (xValues != null) {
                    scatterXValuesByIndex[index] = xValues;
                }
            }

            if (scatterXValuesByIndex.Count == 0) {
                return data;
            }

            var series = new List<ExcelChartSeries>(data.Series.Count);
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                if (!scatterXValuesByIndex.TryGetValue(i, out IReadOnlyList<double>? xValues)) {
                    series.Add(current);
                    continue;
                }

                series.Add(current.WithXValues(xValues));
            }

            return new ExcelChartData(data.Categories, series);
        }

        private static bool TryReadCachedNumberValues(NumberReference? reference, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            values = null;
            NumberingCache? cache = reference?.GetFirstChild<NumberingCache>();
            return TryReadNumberPoints(cache, pointBudget, out values);
        }

        private static bool TryReadNumberLiteralValues(NumberLiteral? literal, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            return TryReadNumberPoints(literal, pointBudget, out values);
        }

        private static bool TryReadNumberPoints(OpenXmlCompositeElement? container, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
            values = null;
            if (container == null) {
                return false;
            }

            uint? pointCount = container.GetFirstChild<PointCount>()?.Val?.Value;
            int actualPointCount = container.Elements<NumericPoint>().Take(MaxChartDataPoints + 1).Count();
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
                foreach (NumericPoint point in container.Elements<NumericPoint>()) {
                    string? text = point.NumericValue?.Text;
                    if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                        return false;
                    }

                    uint index = point.Index?.Value ?? (uint)nextUnindexedIndex++;
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
            foreach (NumericPoint point in container.Elements<NumericPoint>().OrderBy(point => point.Index?.Value ?? uint.MaxValue)) {
                string? text = point.NumericValue?.Text;
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                    return false;
                }

                numericValues.Add(value);
            }

            values = numericValues;
            return numericValues.Count > 0;
        }

        private static bool TryReadCachedStringValues(StringReference? reference, ChartDataPointBudget pointBudget, out IReadOnlyList<string>? values) {
            values = null;
            StringCache? cache = reference?.GetFirstChild<StringCache>();
            if (cache == null) {
                return false;
            }

            int actualPointCount = cache.Elements<StringPoint>().Take(MaxChartDataPoints + 1).Count();
            if (actualPointCount > MaxChartDataPoints || !pointBudget.TryCharge(actualPointCount)) return false;
            string[] cachedValues = cache.Elements<StringPoint>()
                .OrderBy(point => point.Index?.Value ?? uint.MaxValue)
                .Select(point => point.NumericValue?.Text ?? string.Empty)
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

        private static bool TryReadReferencedNumberValues(ExcelSheet contextSheet, string? formula, ChartDataPointBudget pointBudget, out IReadOnlyList<double>? values) {
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
                        numericValues.Add(0D);
                        continue;
                    }

                    if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)) {
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

        private static bool TryGetChartElementType(OpenXmlElement element, out ExcelChartType chartType) {
            switch (element) {
                case BarChart bar:
                    BarDirectionValues direction = bar.GetFirstChild<BarDirection>()?.Val?.Value ?? BarDirectionValues.Column;
                    BarGroupingValues barGrouping = bar.GetFirstChild<BarGrouping>()?.Val?.Value ?? BarGroupingValues.Clustered;
                    if (direction == BarDirectionValues.Bar) {
                        chartType = barGrouping == BarGroupingValues.PercentStacked
                            ? ExcelChartType.BarStacked100
                            : barGrouping == BarGroupingValues.Stacked ? ExcelChartType.BarStacked : ExcelChartType.BarClustered;
                    } else {
                        chartType = barGrouping == BarGroupingValues.PercentStacked
                            ? ExcelChartType.ColumnStacked100
                            : barGrouping == BarGroupingValues.Stacked ? ExcelChartType.ColumnStacked : ExcelChartType.ColumnClustered;
                    }

                    return true;
                case LineChart line:
                    GroupingValues lineGrouping = line.GetFirstChild<Grouping>()?.Val?.Value ?? GroupingValues.Standard;
                    chartType = lineGrouping == GroupingValues.PercentStacked
                        ? ExcelChartType.LineStacked100
                        : lineGrouping == GroupingValues.Stacked ? ExcelChartType.LineStacked : ExcelChartType.Line;
                    return true;
                case AreaChart area:
                    GroupingValues areaGrouping = area.GetFirstChild<Grouping>()?.Val?.Value ?? GroupingValues.Standard;
                    chartType = areaGrouping == GroupingValues.PercentStacked
                        ? ExcelChartType.AreaStacked100
                        : areaGrouping == GroupingValues.Stacked ? ExcelChartType.AreaStacked : ExcelChartType.Area;
                    return true;
                case ScatterChart:
                    chartType = ExcelChartType.Scatter;
                    return true;
                case RadarChart:
                    chartType = ExcelChartType.Radar;
                    return true;
                case PieChart:
                    chartType = ExcelChartType.Pie;
                    return true;
                case Pie3DChart:
                    chartType = ExcelChartType.Pie3D;
                    return true;
                case DoughnutChart:
                    chartType = ExcelChartType.Doughnut;
                    return true;
                default:
                    chartType = default;
                    return false;
            }
        }
    }
}
