using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

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
            ArgumentNullException.ThrowIfNull(categories);
            string[] values = categories.Select(value => value ?? string.Empty).ToArray();
            OfficeXmlNode chart = ResolveChart();
            if (chart.IsNone) return false;

            bool updated = false;
            foreach (OfficeXmlNode series in EnumerateSeries(chart)) {
                OfficeXmlNode categoryData = series.Child(NsC, "cat");
                if (categoryData.IsNone) continue;
                categoryData.RemoveAllChildElements();
                categoryData.Append(BuildStringLiteral(categoryData, values));
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
            ArgumentNullException.ThrowIfNull(values);
            double[] materialized = values.ToArray();
            if (materialized.Any(value => double.IsNaN(value) || double.IsInfinity(value))) {
                throw new ArgumentOutOfRangeException(nameof(values), "Chart series values must be finite numbers.");
            }

            OfficeXmlNode series = FindSeries(seriesIndex);
            if (series.IsNone) return false;

            OfficeXmlNode valueContainer = series.Child(NsC, "val");
            if (valueContainer.IsNone) valueContainer = series.Child(NsC, "yVal");
            if (valueContainer.IsNone) return false;

            string formatCode = GetNumberFormatCode(valueContainer);
            valueContainer.RemoveAllChildElements();
            valueContainer.Append(BuildNumberLiteral(valueContainer, materialized, formatCode));
            return true;
        }

        /// <summary>Replaces the cached name of an existing series with a literal value.</summary>
        /// <param name="seriesIndex">The persisted chart series index (<c>c:idx</c>).</param>
        /// <param name="name">The new series name.</param>
        /// <returns><see langword="true"/> when the indexed series has a mutable series-text element.</returns>
        public bool TrySetSeriesName(uint seriesIndex, string name) {
            ArgumentNullException.ThrowIfNull(name);
            OfficeXmlNode series = FindSeries(seriesIndex);
            OfficeXmlNode seriesText = series.IsNone ? default : series.Child(NsC, "tx");
            if (seriesText.IsNone) return false;

            seriesText.RemoveAllChildElements();
            AddText(seriesText, NsC, "v", name);
            return true;
        }

        private OfficeXmlNode ResolveChart() {
            if (_chart.IsNone && !ChartSpaceNode.IsNone) _chart = ChartSpaceNode.Child(NsC, "chart");
            return _chart;
        }

        private OfficeXmlNode FindSeries(uint seriesIndex) {
            if (ResolveChart().IsNone) return default;
            foreach (OfficeXmlNode series in EnumerateSeries(_chart)) {
                OfficeXmlNode index = series.Child(NsC, "idx");
                if (!index.IsNone && index.UIntVal() == seriesIndex) return series;
            }
            return default;
        }

        /// <summary>
        /// The bar, line, area, radar, pie and scatter series of the chart's plot area, in document order. Series in
        /// extension markup are not plotted and are skipped.
        /// </summary>
        private static IEnumerable<OfficeXmlNode> EnumerateSeries(OfficeXmlNode chart) {
            // Series are children of the chart groups, which are children of the plot area.
            OfficeXmlNode plot = PlotAreaOf(chart);
            if (plot.IsNone) yield break;
            foreach (OfficeXmlNode group in ChartGroups(plot)) {
                if (!IsSupportedSeriesGroup(group)) continue;
                foreach (OfficeXmlNode series in GroupSeries(group)) yield return series;
            }
        }

        private static bool IsSupportedSeriesGroup(OfficeXmlNode group) {
            return group.NamespaceId == NsC && group.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or
                "stockChart" or "areaChart" or "area3DChart" or "radarChart" or "pieChart" or "pie3DChart" or "doughnutChart" or
                "ofPieChart" or "scatterChart";
        }

        private static string GetNumberFormatCode(OfficeXmlNode valueContainer) {
            OfficeXmlNode reference = valueContainer.Child(NsC, "numRef");
            OfficeXmlNode cache = reference.IsNone ? default : reference.Child(NsC, "numCache");
            string? cached = cache.IsNone ? null : cache.ChildText(NsC, "formatCode");
            if (cached != null) return cached;
            OfficeXmlNode literal = valueContainer.Child(NsC, "numLit");
            return (literal.IsNone ? null : literal.ChildText(NsC, "formatCode")) ?? "General";
        }

        private static OfficeXmlNode BuildNumberLiteral(OfficeXmlNode context, IReadOnlyList<double> values, string formatCode) {
            OfficeXmlNode literal = NewC(context, "numLit");
            AddText(literal, NsC, "formatCode", formatCode);
            AddC(literal, "ptCount", (uint)values.Count);
            for (int index = 0; index < values.Count; index++) {
                AddPoint(literal, (uint)index, OfficeIMO.Core.Internal.OfficeRoundTripText.Format(values[index], CultureInfo.InvariantCulture));
            }
            return literal;
        }

        private static OfficeXmlNode BuildStringLiteral(OfficeXmlNode context, IReadOnlyList<string> values) {
            OfficeXmlNode literal = NewC(context, "strLit");
            AddC(literal, "ptCount", (uint)values.Count);
            for (int index = 0; index < values.Count; index++) {
                AddPoint(literal, (uint)index, values[index]);
            }
            return literal;
        }
    }
}
