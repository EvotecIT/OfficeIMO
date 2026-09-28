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
            if (plot == null || plot.GetFirstChild<C.DataTable>() != null ||
                chart.Parent?.ChildElements.Any(element => element.LocalName == "userShapes") == true) return null;
            var layers = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(plot.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)), maximumPoints);
            if (layers.Count == 0) return null;
            ValidatePlotBudget(plot, maximumPoints);
            if (HasUnsupportedBubbleSourceVisibility(part, chart) ||
                chart.Parent?.ChildElements.Any(element => element.LocalName == "style") == true ||
                part.Parts.Any(relationship => relationship.OpenXmlPart.ContentType.IndexOf("chartstyle", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    relationship.OpenXmlPart.ContentType.IndexOf("chartcolorstyle", StringComparison.OrdinalIgnoreCase) >= 0)) return null;
            var axisGroups = OfficeOpenXmlChartAxisGroups.Create(plot);
            if (!HasSupportedProjectionAxisGroups(layers, axisGroups)) return null;
            var projectionBudget = new ProjectionBudget();
            var series = new List<Series>();
            var stackedLayers = new HashSet<(OfficeChartKind Kind, OfficeChartAxisGroup AxisGroup)>();
            IReadOnlyList<string>? categories = null;
            for (int index = 0; index < layers.Count; index++) {
                var layer = layers[index];
                if (layer is not C.BarChart && layer is not C.LineChart && layer is not C.AreaChart && layer is not C.RadarChart &&
                    layer is not C.ScatterChart && layer is not C.BubbleChart && layer is not C.PieChart && layer is not C.DoughnutChart) return null;
                if (!TryReadKind(layer, out var layerKind)) return null;
                // Separate native stacked layers have independent stacks. The flat
                // chart model groups them by kind and axis, so it cannot preserve both.
                if (layerKind is OfficeChartKind.BarStacked or OfficeChartKind.BarStacked100 or
                    OfficeChartKind.ColumnStacked or OfficeChartKind.ColumnStacked100 or
                    OfficeChartKind.LineStacked or OfficeChartKind.LineStacked100 or
                    OfficeChartKind.AreaStacked or OfficeChartKind.AreaStacked100) {
                    if (!stackedLayers.Add((layerKind, axisGroups.Read(layer)))) return null;
                }
                if (HasUnsupportedLayerPresentation(layer)) return null;
                if (HasIncompleteNumericProjectionCaches(layer, maximumPoints)) return null;
                if (index == 0) kind = layerKind;
                Result? data;
                if (layer is C.BubbleChart bubble) {
                    if (layers.Count != 1 || HasUnsupportedBubblePresentation(part, chart, plot, bubble, maximumPointOverrides: 1_000_000)) return null;
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
            if (series.Where(item => item.SourceOrder.HasValue).GroupBy(item => item.SourceOrder).Any(group => group.Count() > 1)) return null;
            series = series.OrderBy(item => item.SourceOrder ?? uint.MaxValue).ToList();
            var legend = chart.GetFirstChild<C.Legend>();
            var hidden = new HashSet<uint>(legend?.Elements<C.LegendEntry>()
                .Where(entry => entry.GetFirstChild<C.Delete>() is C.Delete delete && delete.Val?.Value != false)
                .Select(entry => entry.GetFirstChild<C.Index>()?.Val?.Value)
                .Where(value => value.HasValue).Select(value => value!.Value) ?? Enumerable.Empty<uint>());
            bool categoryLegend = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut;
            series = series.Select((item, index) => new Series(item.SourceIndex,
                item.Data.WithLegendVisibility(legend != null && (categoryLegend ||
                    !hidden.Contains((uint)index))), item.HasUnsupportedAppearance, item.SourceOrder)).ToList();
            var result = new Result(categories, series);
            // The same authoring contract declares supported family/axis combinations.
            OfficeOpenXmlChartWriter.ValidateSharedChartData(result.ToData(), kind);
            return result;
        }

        private static bool HasUnsupportedLayerPresentation(OpenXmlCompositeElement layer) {
            if (layer.Descendants<C.Symbol>().Any(symbol => symbol.Val?.Value == C.MarkerStyleValues.Auto)) return true;
            bool inheritedMarkers = layer is C.LineChart && layer.GetFirstChild<C.ShowMarker>()?.Val?.Value != false ||
                layer is C.RadarChart radarMarkers && radarMarkers.RadarStyle?.Val?.Value == C.RadarStyleValues.Marker ||
                layer is C.ScatterChart scatterMarkers && scatterMarkers.ScatterStyle?.Val?.Value != C.ScatterStyleValues.Line &&
                    scatterMarkers.ScatterStyle?.Val?.Value != C.ScatterStyleValues.Smooth;
            if (inheritedMarkers && layer.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName == "ser").Any(series => series.GetFirstChild<C.Marker>()?.Symbol?.Val == null)) return true;
            if (layer.Descendants().Any(element => element is C.Trendline or C.ErrorBars or C.DropLines or C.HighLowLines or C.UpDownBars or C.SeriesLines)) return true;
            if (layer is C.BarChart && layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(element => element.LocalName == "ser")
                .Any(series => series.Descendants<C.InvertIfNegative>().Any(invert => invert.Val?.Value != false) &&
                    series.GetFirstChild<C.Values>()?.Descendants<C.NumericValue>().Any(value =>
                        double.TryParse(value.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number) && number < 0D) == true)) return true;
            bool radial = layer is C.PieChart or C.DoughnutChart;
            if (radial) {
                // Radial default colouring is per category in the shared renderer.
                if (layer.GetFirstChild<C.VaryColors>()?.Val?.Value != true &&
                    layer.Elements<C.PieChartSeries>().Any(series => series.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.SolidFill>() == null)) return true;
            } else if (layer is not C.BubbleChart && IsVaryColorsEnabled(layer.GetFirstChild<C.VaryColors>())) return true;
            if (layer is C.BarChart bars) {
                if (bars.Descendants().Any(element => element.LocalName == "shape" &&
                    element.GetAttributes().Any(attribute => attribute.LocalName == "val" && attribute.Value != "box"))) return true;
                if (bars.GetFirstChild<C.GapWidth>()?.Val?.Value is ushort gap && gap != 150) return true;
                var grouping = bars.BarGrouping?.Val?.Value;
                int expectedOverlap = grouping == null || grouping == C.BarGroupingValues.Clustered ? 0 : 100;
                if (bars.GetFirstChild<C.Overlap>()?.Val?.Value is sbyte overlap && overlap != expectedOverlap) return true;
            }
            return false;
        }

        private static bool HasSupportedProjectionAxisGroups(System.Collections.Generic.IReadOnlyList<OpenXmlCompositeElement> layers,
            OfficeOpenXmlChartAxisGroups.Groups groups) {
            uint? scatterX = null, scatterY = null;
            foreach (var scatter in layers.OfType<C.ScatterChart>()) {
                var axes = scatter.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                if (axes.Length != 2 || axes[0] is not C.ValueAxis x || axes[1] is not C.ValueAxis y ||
                    ReferenceEquals(x, y) || x.AxisPosition?.Val?.Value != C.AxisPositionValues.Bottom ||
                    y.AxisPosition?.Val?.Value != C.AxisPositionValues.Left) return false;
                uint? xId = x.AxisId?.Val?.Value, yId = y.AxisId?.Val?.Value;
                if (scatterX.HasValue && (scatterX != xId || scatterY != yId)) return false;
                scatterX = xId;
                scatterY = yId;
            }
            var categoryLayers = layers.Where(layer => layer is C.BarChart or C.LineChart or C.AreaChart or C.RadarChart).ToArray();
            if (categoryLayers.Length == 0) return true;
            // The projection contract supports the conventional primary pair followed by
            // a right/top secondary pair. Other native arrangements need independent
            // axis identity and placement metadata instead of guessing from a side.
            if (groups.Read(categoryLayers[0]) != OfficeChartAxisGroup.Primary) return false;
            uint? primaryCategoryId = null, primaryValueId = null;
            uint? secondaryCategoryId = null, secondaryValueId = null;
            foreach (var layer in categoryLayers) {
                var axes = layer.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
                if (axes.Length != 2 || axes.Any(axis => axis == null)) return false;
                var category = axes.SingleOrDefault(axis => axis is C.CategoryAxis);
                var value = axes.SingleOrDefault(axis => axis is C.ValueAxis);
                if (category == null || value == null) return false;
                bool secondary = groups.Read(layer) == OfficeChartAxisGroup.Secondary;
                bool horizontal = layer is C.BarChart bars && bars.BarDirection?.Val?.Value == C.BarDirectionValues.Bar;
                if (secondary && horizontal) return false;
                var expectedCategory = secondary ? C.AxisPositionValues.Top : horizontal ? C.AxisPositionValues.Left : C.AxisPositionValues.Bottom;
                var expectedValue = secondary ? C.AxisPositionValues.Right : horizontal ? C.AxisPositionValues.Bottom : C.AxisPositionValues.Left;
                if (category.GetFirstChild<C.AxisPosition>()?.Val?.Value != expectedCategory || value.GetFirstChild<C.AxisPosition>()?.Val?.Value != expectedValue) return false;
                if (!secondary) {
                    uint? categoryId = category.GetFirstChild<C.AxisId>()?.Val?.Value;
                    uint? valueId = value.GetFirstChild<C.AxisId>()?.Val?.Value;
                    if (primaryCategoryId.HasValue && categoryId != primaryCategoryId) return false;
                    if (primaryValueId.HasValue && valueId != primaryValueId) return false;
                    primaryCategoryId = categoryId;
                    primaryValueId = valueId;
                } else {
                    uint? categoryId = category.GetFirstChild<C.AxisId>()?.Val?.Value;
                    uint? valueId = value.GetFirstChild<C.AxisId>()?.Val?.Value;
                    if (secondaryCategoryId.HasValue && categoryId != secondaryCategoryId) return false;
                    if (secondaryValueId.HasValue && valueId != secondaryValueId) return false;
                    secondaryCategoryId = categoryId;
                    secondaryValueId = valueId;
                }
            }
            return true;
        }

        private static bool HasIncompleteNumericProjectionCaches(OpenXmlCompositeElement layer, int maximumPoints) {
            IReadOnlyList<string>? sharedCategories = null;
            foreach (var series in layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(element => element.LocalName == "ser")) {
                int? seriesLength = null;
                OpenXmlElement?[] caches = layer is C.BubbleChart
                    ? new OpenXmlElement?[] { series.GetFirstChild<C.XValues>(), series.GetFirstChild<C.YValues>(), series.GetFirstChild<C.BubbleSize>() }
                    : layer is C.ScatterChart
                        ? new OpenXmlElement?[] { series.GetFirstChild<C.XValues>(), series.GetFirstChild<C.YValues>() }
                        : new OpenXmlElement?[] { series.GetFirstChild<C.Values>() };
                foreach (var cache in caches) {
                    if (cache == null) return true;
                    var points = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(cache.Descendants<C.NumericPoint>(), maximumPoints);
                    int length = OfficeOpenXmlChartCacheReader.GetCachedPointLength(cache, points, point => point.Index?.Value, maximumPoints);
                    if (length == 0 || points.Count != length || points.Any(point => point.Index?.Value == null) ||
                        points.Select(point => point.Index!.Value).Distinct().Count() != length ||
                        points.Any(point => !double.TryParse(point.NumericValue?.Text, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double number) || double.IsNaN(number) || double.IsInfinity(number))) return true;
                    if (seriesLength.HasValue && seriesLength.Value != length) return true;
                    seriesLength = length;
                }
                if (!seriesLength.HasValue) return true;
                if (layer is C.ScatterChart or C.BubbleChart) continue;
                if (series.GetFirstChild<C.CategoryAxisData>()?.GetFirstChild<C.MultiLevelStringReference>() != null) return true;
                var categories = OfficeOpenXmlChartCacheReader.ReadCachedStrings(series.GetFirstChild<C.CategoryAxisData>(), maximumPoints);
                if (categories.Count != seriesLength.Value ||
                    sharedCategories != null && !sharedCategories.SequenceEqual(categories, StringComparer.Ordinal)) return true;
                sharedCategories = categories;
            }
            return false;
        }
    }
}
