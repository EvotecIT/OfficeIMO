using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static Result? ReadPlot(ChartPart part, C.Chart chart, A.ColorScheme? scheme, int maximumPoints,
            out OfficeChartKind kind, out double bubbleScale, out OfficeChartBubbleSizeMode bubbleMode) {
            kind = default; bubbleScale = 100; bubbleMode = OfficeChartBubbleSizeMode.Area;
            C.PlotArea? plot = chart.PlotArea;
            if (plot == null) return null;
            var layers = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(plot.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)), maximumPoints);
            if (layers.Count == 0) return null;
            ValidatePlotBudget(plot, maximumPoints);
            var axisGroups = OfficeOpenXmlChartAxisGroups.Create(plot);
            var projectionBudget = new ProjectionBudget();
            var series = new List<Series>();
            IReadOnlyList<string>? categories = null;
            for (int index = 0; index < layers.Count; index++) {
                var layer = layers[index];
                if (layer is not C.BarChart && layer is not C.LineChart && layer is not C.AreaChart && layer is not C.RadarChart &&
                    layer is not C.ScatterChart && layer is not C.BubbleChart && layer is not C.PieChart && layer is not C.DoughnutChart) return null;
                if (!TryReadKind(layer, out var layerKind)) return null;
                if (index == 0) kind = layerKind;
                Result? data;
                if (layer is C.BubbleChart bubble) {
                    if (layers.Count != 1 || HasUnsupportedBubblePresentation(part, chart, plot, bubble, maximumPoints)) return null;
                    bubbleScale = bubble.GetFirstChild<C.BubbleScale>()?.Val?.Value ?? 100;
                    bubbleMode = bubble.GetFirstChild<C.SizeRepresents>()?.Val?.Value == C.SizeRepresentsValues.Width
                        ? OfficeChartBubbleSizeMode.Width : OfficeChartBubbleSizeMode.Area;
                    if (bubbleScale > 300) return null;
                    data = ReadBubbles(bubble.Elements<C.BubbleChartSeries>(), scheme, maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget);
                } else if (layer is C.ScatterChart scatter) {
                    if (layers.Count != 1) return null;
                    data = ReadScatter(scatter.Elements<C.ScatterChartSeries>(), scheme, maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget);
                } else {
                    if (layers.Count > 1 && (layer is C.PieChart || layer is C.DoughnutChart || layer is C.RadarChart)) return null;
                    if (layer is C.RadarChart radar && radar.RadarStyle?.Val?.Value == C.RadarStyleValues.Filled) return null;
                    data = ReadCategories(layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser"),
                        layerKind, scheme, axisGroups.Read(layer), maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget);
                }
                if (data == null || data.Series.Any(item => item.HasUnsupportedAppearance)) return null;
                if (layer is C.PieChart && data.Series.Count != 1) return null;
                if (categories != null && !categories.SequenceEqual(data.Categories, StringComparer.Ordinal)) return null;
                categories ??= data.Categories;
                series.AddRange(data.Series);
            }
            if (categories == null || series.Count == 0) return null;
            var legend = chart.GetFirstChild<C.Legend>();
            var hidden = new HashSet<uint>(legend?.Elements<C.LegendEntry>()
                .Where(entry => entry.GetFirstChild<C.Delete>() is C.Delete delete && delete.Val?.Value != false)
                .Select(entry => entry.GetFirstChild<C.Index>()?.Val?.Value)
                .Where(value => value.HasValue).Select(value => value!.Value) ?? Enumerable.Empty<uint>());
            bool categoryLegend = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut;
            bool bubbleLegend = kind == OfficeChartKind.Bubble;
            series = series.Select((item, index) => new Series(item.SourceIndex,
                item.Data.WithLegendVisibility(legend != null && (categoryLegend ||
                    !hidden.Contains(bubbleLegend ? (uint)index : item.SourceIndex))), item.HasUnsupportedAppearance)).ToList();
            var result = new Result(categories, series);
            // The same authoring contract declares supported family/axis combinations.
            OfficeOpenXmlChartWriter.ValidateSharedChartData(result.ToData(), kind);
            return result;
        }
    }
}
