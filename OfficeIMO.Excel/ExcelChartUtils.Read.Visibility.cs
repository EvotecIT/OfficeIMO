using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

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
        internal static ExcelChartData? SelectVisiblePoints(ExcelPartRef chartPart, ExcelSheet contextSheet, ExcelChartData data,
            out bool[]? hiddenSeries) {
            hiddenSeries = null;
            OfficeXmlNode chart = chartPart.Root.Child(C, "chart");
            OfficeXmlNode plotArea = chart.Child(C, "plotArea");
            if (chart.IsNone || plotArea.IsNone) return data;
            OfficeXmlNode bubble = plotArea.Child(C, "bubbleChart");
            IReadOnlyList<OfficeXmlNode> seriesList = bubble.Exists
                ? bubble.Children(C, "ser").ToList().OrderBy(item => item.Child(C, "order").UInt32Val() ?? uint.MaxValue).ToList()
                : GetChartSeries(plotArea);
            if (seriesList.Count != data.Series.Count) return data;

            var hiddenCellsBySheet = new Dictionary<ExcelSheet, OfficeOpenXmlChartHiddenCells>();
            int[]? VisibleOffsets(OfficeXmlNode source, int count) => GetVisibleOffsets(source, count, contextSheet, hiddenCellsBySheet);

            var hidden = new bool[data.Series.Count];
            var series = new IReadOnlyList<double>[data.Series.Count];
            var xValues = new IReadOnlyList<double>?[data.Series.Count];
            var sizes = new IReadOnlyList<double>?[data.Series.Count];
            bool changed = false;
            for (int index = 0; index < data.Series.Count; index++) {
                ExcelChartSeries current = data.Series[index];
                OfficeXmlNode element = seriesList[index];
                OfficeXmlNode valueSource = IsXySeries(element)
                    ? element.Child(C, "yVal")
                    : element.Child(C, "val");
                int[]? valueOffsets = VisibleOffsets(valueSource, current.Values.Count);
                int[]? xOffsets = current.XValues == null ? null : VisibleOffsets(element.Child(C, "xVal"), current.XValues.Count);
                int[]? sizeOffsets = current.BubbleSizes == null ? null : VisibleOffsets(element.Child(C, "bubbleSize"), current.BubbleSizes.Count);
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
            static OfficeXmlNode CategorySource(OfficeXmlNode element) => IsXySeries(element)
                ? element.Child(C, "xVal")
                : element.Child(C, "cat");
            static string? CategoryFormula(OfficeXmlNode source) => ReferenceFormula(source, "numRef") ??
                ReferenceFormula(source, "strRef");
            // Office takes the labels from the first series it still plots. The data holds the first series'
            // labels; another series with its own label cells is read from its cached labels.
            int labelSeries = Math.Max(0, Array.IndexOf(hidden, false));
            OfficeXmlNode categorySource = CategorySource(seriesList[labelSeries]);
            IReadOnlyList<string> sourceCategories = data.Categories;
            if (labelSeries > 0 && !string.Equals(CategoryFormula(categorySource), CategoryFormula(CategorySource(seriesList[0])),
                    StringComparison.OrdinalIgnoreCase)) {
                IReadOnlyList<string>? cached = ReadCachedLabels(categorySource);
                if (cached == null) return null;
                sourceCategories = cached;
            }
            int[]? categoryOffsets = VisibleOffsets(categorySource, sourceCategories.Count);
            // Hidden category cells change the labels even when every value cell is visible.
            if (!changed && categoryOffsets == null) return data;
            IReadOnlyList<string> categories = Select(sourceCategories, categoryOffsets);
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

        /// <summary>Keeps the custom error amounts of the points a visible-cells-only chart plots, as
        /// <see cref="SelectVisiblePoints(ExcelPartRef, ExcelSheet, ExcelChartData, out bool[])"/> keeps their values.</summary>
        internal static Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>> CreateVisibleNodeAmountSelector(
            ExcelSheet contextSheet) {
            var hiddenCellsBySheet = new Dictionary<ExcelSheet, OfficeOpenXmlChartHiddenCells>();
            return (source, amounts) => Select(amounts, GetVisibleOffsets(source, amounts.Count, contextSheet, hiddenCellsBySheet));
        }

        // The visible cells of a source range, as offsets into its cached points; null when the range does not map one
        // cell to each cached point or nothing in it is hidden.
        private static int[]? GetVisibleOffsets(OfficeXmlNode source, int count, ExcelSheet contextSheet,
            Dictionary<ExcelSheet, OfficeOpenXmlChartHiddenCells> hiddenCellsBySheet) {
            string? formula = ReferenceFormula(source, "numRef") ??
                ReferenceFormula(source, "strRef");
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
                hidden = ExcelPartRef.CreateHiddenCells(sheet);
                hiddenCellsBySheet.Add(sheet, hidden);
            }
            return hidden.GetVisibleOffsets((uint)firstRow, (uint)lastRow, (uint)firstColumn, (uint)lastColumn);
        }

        private static IReadOnlyList<T> Select<T>(IReadOnlyList<T> values, int[]? offsets) =>
            offsets == null ? values : offsets.Where(offset => offset < values.Count).Select(offset => values[offset]).ToArray();

        // The cached labels of a category or X range, in point order; null when the cache does not cover every point.
        private static IReadOnlyList<string>? ReadCachedLabels(OfficeXmlNode source) {
            OfficeXmlNode cache = FirstCache(source);
            uint? count = cache.Child(C, "ptCount").UInt32Val();
            if (cache.IsNone || count == null || count.Value > 1_000_000U) return null;
            var labels = new string?[count.Value];
            foreach (OfficeXmlNode point in cache.Children(C, "pt")) {
                uint? index = point.UInt32Attribute("idx");
                if (index == null || index.Value >= count.Value) return null;
                labels[index.Value] = point.ChildText(C, "v") ?? string.Empty;
            }
            return labels.Any(label => label == null) ? null : labels.Select(label => label!).ToArray();
        }

        // The first c:strCache or c:numCache under a source, in document order and outside extension lists, or none.
        private static OfficeXmlNode FirstCache(OfficeXmlNode element) =>
            element.FirstDescendantOutsideExtensions(child => child.Is(C, "strCache") || child.Is(C, "numCache"));

        /// <summary>Drops the series <see cref="SelectVisiblePoints(ExcelPartRef, ExcelSheet, ExcelChartData, out bool[])"/> found hidden, after formatting has been matched to them.</summary>
        internal static ExcelChartData RemoveSeries(ExcelChartData data, bool[]? hiddenSeries) =>
            hiddenSeries == null || hiddenSeries.Length != data.Series.Count
                ? data
                : new ExcelChartData(data.Categories, data.Series.Where((_, index) => !hiddenSeries[index]));
    }
}
