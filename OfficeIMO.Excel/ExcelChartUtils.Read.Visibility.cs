using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        /// <summary>Leaves out the points whose source cells are hidden, as Office plots a chart that shows
        /// visible cells only (c:plotVisOnly).</summary>
        /// <remarks>Each referenced range loses its hidden cells and closes the gap, so series with differently
        /// placed ranges can end up with different lengths; the shorter ones are padded with missing values and
        /// the categories with empty labels. A series whose value cells are all hidden plots nothing: it stays in
        /// the returned data, so per-series formatting still lines up with the chart XML, and is listed in
        /// <paramref name="hiddenSeries"/> for <see cref="RemoveSeries"/>. A chart whose every series is hidden
        /// plots nothing, so all of them are listed.</remarks>
        /// <returns>The visible data, or null when a zero blank rule would draw the padding.</returns>
        internal static ExcelChartData? SelectVisiblePoints(ChartPart chartPart, ExcelSheet contextSheet, ExcelChartData data,
            out bool[]? hiddenSeries) {
            hiddenSeries = null;
            Chart? chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            PlotArea? plotArea = chart?.GetFirstChild<PlotArea>();
            if (chart == null || plotArea == null) return data;
            IReadOnlyList<OpenXmlCompositeElement> seriesList = plotArea.GetFirstChild<BubbleChart>() is BubbleChart bubble
                ? bubble.Elements<BubbleChartSeries>().OrderBy(item => item.GetFirstChild<Order>()?.Val?.Value ?? uint.MaxValue).ToList()
                : GetChartSeries(plotArea);
            if (seriesList.Count != data.Series.Count) return data;

            var hiddenCellsBySheet = new Dictionary<ExcelSheet, OfficeOpenXmlChartHiddenCells>();
            int[]? VisibleOffsets(OpenXmlElement? source, int count) {
                string? formula = source?.GetFirstChild<NumberReference>()?.Formula?.Text ??
                    source?.GetFirstChild<StringReference>()?.Formula?.Text;
                if (!TryParseSheetQualifiedRange(formula, out string sheetName, out string range) ||
                    !TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) ||
                    (long)(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) != count) return null;
                ExcelSheet sheet;
                try {
                    sheet = contextSheet.Document[sheetName];
                } catch {
                    return null;
                }
                if (!hiddenCellsBySheet.TryGetValue(sheet, out OfficeOpenXmlChartHiddenCells? hidden)) {
                    hidden = new OfficeOpenXmlChartHiddenCells(sheet.WorksheetPart.Worksheet!);
                    hiddenCellsBySheet.Add(sheet, hidden);
                }
                return hidden.GetVisibleOffsets((uint)firstRow, (uint)lastRow, (uint)firstColumn, (uint)lastColumn);
            }
            static IReadOnlyList<T> Select<T>(IReadOnlyList<T> values, int[]? offsets) =>
                offsets == null ? values : offsets.Where(offset => offset < values.Count).Select(offset => values[offset]).ToArray();

            var hidden = new bool[data.Series.Count];
            var series = new IReadOnlyList<double>[data.Series.Count];
            var xValues = new IReadOnlyList<double>?[data.Series.Count];
            var sizes = new IReadOnlyList<double>?[data.Series.Count];
            bool changed = false;
            for (int index = 0; index < data.Series.Count; index++) {
                ExcelChartSeries current = data.Series[index];
                OpenXmlCompositeElement element = seriesList[index];
                OpenXmlElement? valueSource = element is ScatterChartSeries or BubbleChartSeries
                    ? element.GetFirstChild<YValues>()
                    : element.GetFirstChild<Values>();
                int[]? valueOffsets = VisibleOffsets(valueSource, current.Values.Count);
                int[]? xOffsets = current.XValues == null ? null : VisibleOffsets(element.GetFirstChild<XValues>(), current.XValues.Count);
                int[]? sizeOffsets = current.BubbleSizes == null ? null : VisibleOffsets(element.GetFirstChild<BubbleSize>(), current.BubbleSizes.Count);
                changed |= valueOffsets != null || xOffsets != null || sizeOffsets != null;
                series[index] = Select(current.Values, valueOffsets);
                xValues[index] = current.XValues == null ? null : Select(current.XValues, xOffsets);
                sizes[index] = current.BubbleSizes == null ? null : Select(current.BubbleSizes, sizeOffsets);
                hidden[index] = series[index].Count == 0;
                if (xValues[index] is IReadOnlyList<double> x && x.Count != series[index].Count) {
                    // Office pairs the visible X and Y cells in order; positions past the shorter range have no point.
                    int count = Math.Min(x.Count, series[index].Count);
                    xValues[index] = x.Take(count).ToArray();
                    series[index] = series[index].Take(count).ToArray();
                    if (sizes[index] != null) return null;
                }
                if (sizes[index] is IReadOnlyList<double> bubbleSizes && bubbleSizes.Count != series[index].Count) return null;
            }
            if (!changed) return data;

            OpenXmlCompositeElement first = seriesList[0];
            OpenXmlElement? categorySource = first is ScatterChartSeries or BubbleChartSeries
                ? first.GetFirstChild<XValues>()
                : first.GetFirstChild<CategoryAxisData>();
            IReadOnlyList<string> categories = Select(data.Categories, VisibleOffsets(categorySource, data.Categories.Count));
            int length = Math.Max(categories.Count, Enumerable.Range(0, series.Length)
                .Where(index => xValues[index] == null && !hidden[index]).Select(index => series[index].Count).DefaultIfEmpty(0).Max());
            bool padsValues = false;
            var visibleSeries = new List<ExcelChartSeries>(data.Series.Count);
            for (int index = 0; index < data.Series.Count; index++) {
                IReadOnlyList<double> values = series[index];
                if (xValues[index] == null && values.Count != length) {
                    padsValues |= !hidden[index];
                    values = values.Concat(Enumerable.Repeat(double.NaN, length - values.Count)).ToArray();
                }
                ExcelChartSeries current = data.Series[index].WithValuesAndXValues(values, xValues[index]);
                visibleSeries.Add(sizes[index] is IReadOnlyList<double> bubbleSizes ? current.WithBubbleSizes(bubbleSizes) : current);
            }
            // Office draws no point past the end of a shortened series, which the padding only matches when
            // missing values are not drawn as zero.
            if (padsValues && OfficeOpenXmlChartPlotSemantics.ReadBlankDisplay(chart) == OfficeIMO.Drawing.OfficeChartBlankDisplay.Zero) return null;
            if (categories.Count < length) categories = categories.Concat(Enumerable.Repeat(string.Empty, length - categories.Count)).ToArray();
            hiddenSeries = hidden.Any(item => item) ? hidden : null;
            return new ExcelChartData(categories, visibleSeries);
        }

        /// <summary>Drops the series <see cref="SelectVisiblePoints"/> found hidden, after formatting has been matched to them.</summary>
        internal static ExcelChartData RemoveSeries(ExcelChartData data, bool[]? hiddenSeries) =>
            hiddenSeries == null || hiddenSeries.Length != data.Series.Count
                ? data
                : new ExcelChartData(data.Categories, data.Series.Where((_, index) => !hiddenSeries[index]));
    }
}
