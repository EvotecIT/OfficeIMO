using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        // Bubble series have independent X/Y/size ranges, not a rectangular category/value range.
        private static ExcelChartData? TryReadBubbleChartData(PlotArea plot, BubbleChart bubble,
            ExcelSheet sheet, ChartDataPointBudget budget) {
            if (plot.Elements<BubbleChart>().Count() != 1 || plot.ChildElements.Any(element =>
                element != bubble && TryGetChartElementType(element, out _))) return null;

            var source = bubble.Elements<BubbleChartSeries>().Take(MaxChartDataPoints + 1).ToList();
            if (source.Count == 0 || source.Count > MaxChartDataPoints) return null;
            var orders = new HashSet<uint>();
            foreach (var item in source) {
                uint? order = item.GetFirstChild<Order>()?.Val?.Value;
                if (!order.HasValue || !orders.Add(order.Value)) return null;
            }

            var series = new List<ExcelChartSeries>(source.Count);
            IReadOnlyList<string>? categories = null;
            foreach (BubbleChartSeries item in source.OrderBy(item => item.GetFirstChild<Order>()!.Val!.Value)) {
                if (!TryReadBubbleNumbers(item.GetFirstChild<XValues>(), sheet, budget, out var x) ||
                    !TryReadBubbleNumbers(item.GetFirstChild<YValues>(), sheet, budget, out var y) ||
                    !TryReadBubbleNumbers(item.GetFirstChild<BubbleSize>(), sheet, budget, out var sizes) ||
                    x!.Count != y!.Count || x.Count != sizes!.Count) return null;

                categories ??= x.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray();
                string name = TryReadSeriesName(item, sheet, out string? resolved) && !string.IsNullOrWhiteSpace(resolved)
                    ? resolved! : "Series " + (series.Count + 1).ToString(CultureInfo.InvariantCulture);
                series.Add(new ExcelChartSeries(name, y, x, OfficeChartType.Bubble, GetChartAxisGroup(plot, bubble))
                    .WithImageExportStyle(null, null, null, null, true, false, null, null, null, null)
                    .WithBubbleSizes(sizes));
            }
            return categories == null ? null : new ExcelChartData(categories, series);
        }

        private static bool TryReadBubbleNumbers(OpenXmlCompositeElement? container, ExcelSheet sheet,
            ChartDataPointBudget budget, out IReadOnlyList<double>? values) {
            values = null;
            if (container == null) return false;
            NumberReference? reference = container.GetFirstChild<NumberReference>();
            if (TryReadReferencedNumberValues(sheet, reference?.Formula?.Text, budget, out values))
                return values!.All(value => !double.IsNaN(value) && !double.IsInfinity(value));

            OpenXmlCompositeElement? cache = container.GetFirstChild<NumberLiteral>() ??
                (OpenXmlCompositeElement?)reference?.GetFirstChild<NumberingCache>();
            if (cache == null) return false;
            var points = cache.Elements<NumericPoint>().Take(MaxChartDataPoints + 1).ToList();
            uint count = cache.GetFirstChild<PointCount>()?.Val?.Value ?? (uint)points.Count;
            if (count == 0 || count != points.Count || !budget.TryCharge(count)) return false;
            var parsed = new double[points.Count];
            var seen = new bool[points.Count];
            foreach (NumericPoint point in points) {
                uint? index = point.Index?.Value;
                if (!index.HasValue || index.Value >= count || seen[index.Value] ||
                    !double.TryParse(point.NumericValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.IsNaN(value) || double.IsInfinity(value)) return false;
                seen[index.Value] = true;
                parsed[index.Value] = value;
            }
            values = parsed;
            return true;
        }
    }
}
