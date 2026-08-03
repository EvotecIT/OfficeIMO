using System.Globalization;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        /// <summary>
        /// Replaces the cached categories on every category-based series with literal values.
        /// Existing worksheet references are detached so Word cannot silently restore stale linked cache values.
        /// Scatter charts do not have category-axis data and are not changed.
        /// </summary>
        /// <param name="categories">The literal category labels to persist.</param>
        /// <returns><see langword="true"/> when at least one imported or code-created series was updated.</returns>
        public bool TrySetCategories(IEnumerable<string> categories) {
            if (categories == null) throw new ArgumentNullException(nameof(categories));
            string[] values = categories.Select(value => value ?? string.Empty).ToArray();
            C.Chart? chart = ResolveChart();
            if (chart == null) return false;

            bool updated = false;
            foreach (OpenXmlCompositeElement series in EnumerateSeries(chart)) {
                C.CategoryAxisData? categoryData = series.GetFirstChild<C.CategoryAxisData>();
                if (categoryData == null) continue;
                categoryData.RemoveAllChildren();
                categoryData.Append(BuildStringLiteral(values));
                updated = true;
            }

            if (updated) Categories = values.ToList();
            return updated;
        }

        /// <summary>
        /// Replaces one series' cached numeric values with literal values. Worksheet references for that value
        /// range are detached, while unrelated embedded-workbook parts are preserved.
        /// </summary>
        /// <param name="seriesIndex">The persisted chart series index (<c>c:idx</c>).</param>
        /// <param name="values">The literal numeric values to persist.</param>
        /// <returns><see langword="true"/> when the indexed series and its numeric value container were found.</returns>
        public bool TrySetSeriesValues(uint seriesIndex, IEnumerable<double> values) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            double[] materialized = values.ToArray();
            if (materialized.Any(value => double.IsNaN(value) || double.IsInfinity(value))) {
                throw new ArgumentOutOfRangeException(nameof(values), "Chart series values must be finite numbers.");
            }

            OpenXmlCompositeElement? series = FindSeries(seriesIndex);
            if (series == null) return false;

            OpenXmlCompositeElement? valueContainer = series.GetFirstChild<C.Values>();
            valueContainer ??= series.GetFirstChild<C.YValues>();
            if (valueContainer == null) return false;

            valueContainer.RemoveAllChildren();
            valueContainer.Append(BuildNumberLiteral(materialized));
            return true;
        }

        /// <summary>Replaces the cached name of an existing series with a literal value.</summary>
        /// <param name="seriesIndex">The persisted chart series index (<c>c:idx</c>).</param>
        /// <param name="name">The new series name.</param>
        /// <returns><see langword="true"/> when the indexed series has a mutable series-text element.</returns>
        public bool TrySetSeriesName(uint seriesIndex, string name) {
            if (name == null) throw new ArgumentNullException(nameof(name));
            OpenXmlCompositeElement? series = FindSeries(seriesIndex);
            C.SeriesText? seriesText = series?.GetFirstChild<C.SeriesText>();
            if (seriesText == null) return false;

            seriesText.RemoveAllChildren();
            seriesText.Append(AddSeries(0U, name));
            return true;
        }

        private C.Chart? ResolveChart() {
            _chart ??= _chartPart?.ChartSpace?.GetFirstChild<C.Chart>();
            return _chart;
        }

        private OpenXmlCompositeElement? FindSeries(uint seriesIndex) => ResolveChart() == null
            ? null
            : EnumerateSeries(_chart!).FirstOrDefault(series =>
                series.GetFirstChild<C.Index>()?.Val?.Value == seriesIndex);

        private static IEnumerable<OpenXmlCompositeElement> EnumerateSeries(C.Chart chart) =>
            chart.PlotArea?.Descendants<OpenXmlCompositeElement>().Where(IsSupportedSeriesElement)
            ?? Enumerable.Empty<OpenXmlCompositeElement>();

        private static bool IsSupportedSeriesElement(OpenXmlCompositeElement element) =>
            element is C.BarChartSeries or
            C.LineChartSeries or
            C.AreaChartSeries or
            C.RadarChartSeries or
            C.PieChartSeries or
            C.ScatterChartSeries;

        private static C.NumberLiteral BuildNumberLiteral(IReadOnlyList<double> values) {
            var literal = new C.NumberLiteral(
                new C.FormatCode { Text = "General" },
                new C.PointCount { Val = (uint)values.Count });
            for (int index = 0; index < values.Count; index++) {
                literal.Append(new C.NumericPoint(
                    new C.NumericValue(values[index].ToString("R", CultureInfo.InvariantCulture))) {
                    Index = (uint)index
                });
            }
            return literal;
        }

        private static C.StringLiteral BuildStringLiteral(IReadOnlyList<string> values) {
            var literal = new C.StringLiteral(new C.PointCount { Val = (uint)values.Count });
            for (int index = 0; index < values.Count; index++) {
                literal.Append(new C.StringPoint(
                    new C.NumericValue(values[index])) {
                    Index = (uint)index
                });
            }
            return literal;
        }
    }
}
