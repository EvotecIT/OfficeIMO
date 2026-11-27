using System.Globalization;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        // Bubble series have independent X/Y/size ranges, not a rectangular category/value range.
        private static ExcelChartData? TryReadBubbleChartData(OfficeXmlNode plot, OfficeXmlNode bubble,
            ExcelSheet sheet, ChartDataPointBudget budget) {
            if (plot.Children(C, "bubbleChart").Count != 1) return null;
            foreach (OfficeXmlNode element in plot.Children()) {
                if (element != bubble && TryGetChartElementType(element, out _)) return null;
            }

            var source = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode item in bubble.Children(C, "ser")) {
                source.Add(item);
                if (source.Count > MaxChartDataPoints) break;
            }
            if (source.Count == 0 || source.Count > MaxChartDataPoints) return null;
            var orders = new HashSet<uint>();
            foreach (var item in source) {
                uint? order = item.Child(C, "order").UInt32Val();
                if (!order.HasValue || !orders.Add(order.Value)) return null;
            }

            var series = new List<ExcelChartSeries>(source.Count);
            OfficeChartAxisGroup axisGroup = GetChartAxisGroup(plot, bubble);
            IReadOnlyList<string>? categories = null;
            foreach (OfficeXmlNode item in source.OrderBy(item => item.Child(C, "order").UInt32Val()!.Value)) {
                if (!TryReadBubbleNumbers(item.Child(C, "xVal"), sheet, budget, out var x) ||
                    !TryReadBubbleNumbers(item.Child(C, "yVal"), sheet, budget, out var y) ||
                    !TryReadBubbleNumbers(item.Child(C, "bubbleSize"), sheet, budget, out var sizes) ||
                    x!.Count != y!.Count || x.Count != sizes!.Count) return null;

                categories ??= x.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                string name = TryReadSeriesName(item, sheet, out string? resolved) && !string.IsNullOrWhiteSpace(resolved)
                    ? resolved! : "Series " + (series.Count + 1).ToString(CultureInfo.InvariantCulture);
                series.Add(new ExcelChartSeries(name, y, x, OfficeChartType.Bubble, axisGroup)
                    .WithImageExportStyle(null, null, null, null, true, false, null, null, null, null)
                    .WithBubbleSizes(sizes));
            }
            return categories == null ? null : new ExcelChartData(categories, series);
        }

        private static bool TryReadBubbleNumbers(OfficeXmlNode container, ExcelSheet sheet,
            ChartDataPointBudget budget, out IReadOnlyList<double>? values) {
            values = null;
            if (container.IsNone) return false;
            OfficeXmlNode reference = container.Child(C, "numRef");
            if (TryReadReferencedNumberValues(sheet, reference.ChildText(C, "f"), budget, out values))
                return values!.All(value => !double.IsNaN(value) && !double.IsInfinity(value));

            OfficeXmlNode cache = container.Child(C, "numLit");
            if (cache.IsNone) cache = reference.Child(C, "numCache");
            if (cache.IsNone) return false;
            var points = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode point in cache.Children(C, "pt")) {
                points.Add(point);
                if (points.Count > MaxChartDataPoints) break;
            }
            uint count = cache.Child(C, "ptCount").UInt32Val() ?? (uint)points.Count;
            if (count == 0 || count != points.Count || !budget.TryCharge(count)) return false;
            var parsed = new double[points.Count];
            var seen = new bool[points.Count];
            foreach (OfficeXmlNode point in points) {
                uint? index = point.UInt32Attribute("idx");
                if (!index.HasValue || index.Value >= count || seen[index.Value] ||
                    !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(point.ChildText(C, "v"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.IsNaN(value) || double.IsInfinity(value)) return false;
                seen[index.Value] = true;
                parsed[index.Value] = value;
            }
            values = parsed;
            return true;
        }
    }
}
