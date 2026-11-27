using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static Result? ReadPlot(IOfficeChartPart part, OfficeXmlNode chart, OfficeXmlNode scheme, int maximumPoints,
            out OfficeChartType kind, out double bubbleScale, out OfficeChartBubbleSizeMode bubbleMode) {
            kind = default; bubbleScale = 100; bubbleMode = OfficeChartBubbleSizeMode.Area;
            OfficeXmlNode plot = chart.Child(NsC, "plotArea");
            if (plot.IsNone || !plot.Child(NsC, "dTable").IsNone ||
                chart.Parent.Children().ToList().Any(element => element.LocalName == "userShapes")) return null;
            var layers = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(plot.Children().ToList()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)), maximumPoints);
            if (layers.Count == 0) return null;
            ValidatePlotBudget(plot, maximumPoints);
            // The classic c:style number only supplies automatic colours, which the series reader
            // resolves. Office does not consult the cs:chartStyle and cs:colorStyle parts when it
            // renders a chart; they only seed later restyling in the Office UI.
            if (!TryReadSourceVisibility(part, chart, out SourceVisibility? visibility, out _)) return null;
            var axisGroups = OfficeOpenXmlChartAxisGroups.Create(plot);
            if (!HasSupportedProjectionAxisGroups(layers, axisGroups)) return null;
            var barGeometry = layers.Where(layer => layer.IsChart("barChart")).Select(bars => {
                string? grouping = bars.ChildVal(NsC, "grouping");
                return (
                    Gap: bars.Child(NsC, "gapWidth").UShortVal() ?? 150,
                    Overlap: grouping == null || grouping == "clustered"
                        ? bars.Child(NsC, "overlap").SByteVal() ?? 0 : 0);
            }).Distinct().Take(2).Count();
            if (barGeometry > 1) return null;
            var projectionBudget = new ProjectionBudget();
            var series = new List<Series>();
            var stackedLayers = new HashSet<(OfficeChartType Kind, OfficeChartAxisGroup AxisGroup)>();
            IReadOnlyList<string>? categories = null;
            for (int index = 0; index < layers.Count; index++) {
                var layer = layers[index];
                if (layer.NamespaceId != NsC || layer.LocalName is not ("barChart" or "lineChart" or "areaChart" or "radarChart" or
                    "scatterChart" or "bubbleChart" or "pieChart" or "doughnutChart" or "stockChart")) return null;
                if (!TryReadKind(layer, out var layerKind)) return null;
                // Separate native stacked layers have independent stacks. The flat
                // chart model groups them by kind and axis, so it cannot preserve both.
                if (layerKind is OfficeChartType.BarStacked or OfficeChartType.BarStacked100 or
                    OfficeChartType.ColumnStacked or OfficeChartType.ColumnStacked100 or
                    OfficeChartType.LineStacked or OfficeChartType.LineStacked100 or
                    OfficeChartType.AreaStacked or OfficeChartType.AreaStacked100) {
                    if (!stackedLayers.Add((layerKind, axisGroups.Read(layer)))) return null;
                }
                if (HasUnsupportedLayerPresentation(layer)) return null;
                if (HasIncompleteNumericProjectionCaches(layer, maximumPoints)) return null;
                if (index == 0) kind = layerKind;
                Result? data;
                if (layer.IsChart("bubbleChart")) {
                    OfficeXmlNode bubble = layer;
                    if (layers.Count != 1 || HasUnsupportedBubblePresentation(part, chart, plot, bubble, maximumPointOverrides: 1_000_000)) return null;
                    bubbleScale = bubble.Child(NsC, "bubbleScale").UIntVal() ?? 100;
                    bubbleMode = bubble.ChildVal(NsC, "sizeRepresents") == "w"
                        ? OfficeChartBubbleSizeMode.Width : OfficeChartBubbleSizeMode.Area;
                    if (bubbleScale > 300) return null;
                    data = ReadBubbles(bubble.Children(NsC, "ser").ToList(), scheme, maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget, visibility: visibility);
                } else if (layer.IsChart("scatterChart")) {
                    if (layers.Count != 1) return null;
                    data = ReadScatter(layer.Children(NsC, "ser").ToList(), scheme, maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget, keepMissingValues: true, visibility: visibility);
                } else {
                    if (layers.Count > 1 && (layer.IsChart("pieChart") || layer.IsChart("doughnutChart") || layer.IsChart("radarChart"))) return null;
                    if (layer.IsChart("radarChart") && layer.ChildVal(NsC, "radarStyle") == "filled") return null;
                    data = ReadCategories(layer.Children().ToList().Where(item => item.LocalName == "ser"),
                        layerKind, scheme, axisGroups.Read(layer), maximumPoints,
                        validatePlot: false, projectionBudget: projectionBudget, keepMissingValues: true, visibility: visibility);
                }
                if (data == null) return null;
                if (data.Series.Any(item => item.HasUnsupportedAppearance)) {
                    if (!OfficeOpenXmlChartFidelityScope.AllowsApproximation)
                        throw new NotSupportedException(data.Series.First(item => item.HasUnsupportedAppearance).UnsupportedAppearance);
                    // Best effort keeps the series' resolved colour and drops the formatting it cannot draw.
                    foreach (var item in data.Series.Where(item => item.HasUnsupportedAppearance))
                        OfficeOpenXmlChartFidelityScope.Record(item.UnsupportedAppearance!.TrimEnd('.') + "; it is drawn in its plain series colour.");
                }
                if (layer.IsChart("pieChart") && data.Series.Count != 1) return null;
                if (categories != null && !categories.SequenceEqual(data.Categories, StringComparer.Ordinal)) return null;
                categories ??= data.Categories;
                series.AddRange(data.Series);
            }
            if (categories == null || series.Count == 0) return null;
            if (series.Where(item => item.SourceOrder.HasValue).GroupBy(item => item.SourceOrder).Any(group => group.Count() > 1)) return null;
            series = series.OrderBy(item => item.SourceOrder ?? uint.MaxValue).ToList();
            var legend = chart.Child(NsC, "legend");
            var hidden = new HashSet<uint>(legend.Children(NsC, "legendEntry").ToList()
                .Where(entry => entry.Child(NsC, "delete") is { IsNone: false } delete && delete.BoolVal() != false)
                .Select(entry => entry.Child(NsC, "idx").UIntVal())
                .Where(value => value.HasValue).Select(value => value!.Value));
            bool categoryLegend = kind == OfficeChartType.Pie || kind == OfficeChartType.Doughnut ||
                !OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plot).IsNone;
            series = series.Select((item, index) => new Series(item.SourceIndex,
                item.Data.WithLegendVisibility(!legend.IsNone && (categoryLegend ||
                    !hidden.Contains((uint)index))), item.UnsupportedAppearance, item.SourceOrder,
                item.HasAutomaticColor)).ToList();
            var result = new Result(categories, series);

            // The same authoring contract declares supported family/axis combinations.
            OfficeOpenXmlChartWriter.ValidateSharedChartData(result.ToData(), kind, allowEmptyValues: true);
            return result;
        }

        private static bool HasUnsupportedLayerPresentation(OfficeXmlNode layer) {
            // Series trendlines and error bars are read with their series.
            bool lineLayer = layer.IsChart("lineChart") || layer.IsChart("stockChart");
            if (layer.AllDescendants().Any(element => element.NamespaceId == NsC && (element.LocalName is "dropLines" or "upDownBars" or "serLines" ||
                element.LocalName == "hiLowLines" && !lineLayer))) {
                if (!OfficeOpenXmlChartFidelityScope.AllowsApproximation) return true;
                OfficeOpenXmlChartFidelityScope.Record("Drop lines, up/down bars and series lines are left out.");
            }
            if (layer.IsChart("barChart") && layer.Children().ToList().Where(element => element.LocalName == "ser")
                .Any(series => series.Descendants(NsC, "invertIfNegative").Any(invert => invert.BoolVal() != false) &&
                    series.Child(NsC, "val").Descendants(NsC, "v").Any(value =>
                        OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value.InnerText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number) && number < 0D))) return true;
            bool radial = layer.IsChart("pieChart") || layer.IsChart("doughnutChart");
            if (radial) {
                // Radial default colouring is per category in the shared renderer.
                if (layer.Child(NsC, "varyColors").BoolVal() != true &&
                    layer.Children(NsC, "ser").ToList().Any(series => series.Child(NsC, "spPr").Child(NsA, "solidFill").IsNone)) return true;
            }
            if (layer.IsChart("barChart")) {
                OfficeXmlNode bars = layer;
                if (bars.AllDescendants().Any(element => element.LocalName == "shape" &&
                    element.Attributes().Any(attribute => attribute.LocalName == "val" && attribute.Value != "box"))) return true;
                if (bars.Child(NsC, "gapWidth").UShortVal() is ushort gap && gap > 500) return true;
                if (bars.Child(NsC, "overlap").SByteVal() is sbyte overlap && (overlap < -100 || overlap > 100)) return true;
            }
            return false;
        }

        private static bool HasSupportedProjectionAxisGroups(System.Collections.Generic.IReadOnlyList<OfficeXmlNode> layers,
            OfficeOpenXmlChartAxisGroups.Groups groups) {
            uint? scatterX = null, scatterY = null;
            foreach (var scatter in layers.Where(layer => layer.IsChart("scatterChart"))) {
                var axes = scatter.Children(NsC, "axId").ToList().Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
                if (axes.Length != 2 || !axes[0].IsChart("valAx") || !axes[1].IsChart("valAx")) return false;
                OfficeXmlNode x = axes[0], y = axes[1];
                if (x == y || x.ChildVal(NsC, "axPos") != "b" ||
                    y.ChildVal(NsC, "axPos") != "l") return false;
                uint? xId = x.Child(NsC, "axId").UIntVal(), yId = y.Child(NsC, "axId").UIntVal();
                if (scatterX.HasValue && (scatterX != xId || scatterY != yId)) return false;
                scatterX = xId;
                scatterY = yId;
            }
            var categoryLayers = layers.Where(layer => layer.NamespaceId == NsC && layer.LocalName is "barChart" or "lineChart" or "areaChart" or "radarChart").ToArray();
            if (categoryLayers.Length == 0) return true;
            // The projection contract supports the conventional primary pair followed by
            // a right/top secondary pair. Other native arrangements need independent
            // axis identity and placement metadata instead of guessing from a side.
            if (groups.Read(categoryLayers[0]) != OfficeChartAxisGroup.Primary) return false;
            uint? primaryCategoryId = null, primaryValueId = null;
            uint? secondaryCategoryId = null, secondaryValueId = null;
            foreach (var layer in categoryLayers) {
                var axes = layer.Children(NsC, "axId").ToList().Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
                if (axes.Length != 2 || axes.Any(axis => axis.IsNone)) return false;
                var category = axes.SingleOrDefault(axis => axis.IsChart("catAx"));
                var value = axes.SingleOrDefault(axis => axis.IsChart("valAx"));
                if (category.IsNone || value.IsNone) return false;
                bool secondary = groups.Read(layer) == OfficeChartAxisGroup.Secondary;
                bool horizontal = layer.IsChart("barChart") && layer.ChildVal(NsC, "barDir") == "bar";
                if (secondary && horizontal) return false;
                string expectedCategory = secondary ? "t" : horizontal ? "l" : "b";
                string expectedValue = secondary ? "r" : horizontal ? "b" : "l";
                // Office writes the hidden secondary category axis of a combo chart at the bottom; its side is not drawn.
                bool hiddenSecondaryCategory = secondary && IsDeletedAxis(category);
                if (category.ChildVal(NsC, "axPos") != expectedCategory && !hiddenSecondaryCategory ||
                    value.ChildVal(NsC, "axPos") != expectedValue) return false;
                if (!secondary) {
                    uint? categoryId = category.Child(NsC, "axId").UIntVal();
                    uint? valueId = value.Child(NsC, "axId").UIntVal();
                    if (primaryCategoryId.HasValue && categoryId != primaryCategoryId) return false;
                    if (primaryValueId.HasValue && valueId != primaryValueId) return false;
                    primaryCategoryId = categoryId;
                    primaryValueId = valueId;
                } else {
                    uint? categoryId = category.Child(NsC, "axId").UIntVal();
                    uint? valueId = value.Child(NsC, "axId").UIntVal();
                    if (secondaryCategoryId.HasValue && categoryId != secondaryCategoryId) return false;
                    if (secondaryValueId.HasValue && valueId != secondaryValueId) return false;
                    secondaryCategoryId = categoryId;
                    secondaryValueId = valueId;
                }
            }
            return true;
        }

        private static bool HasIncompleteNumericProjectionCaches(OfficeXmlNode layer, int maximumPoints) {
            IReadOnlyList<string>? sharedCategories = null;
            bool bubble = layer.IsChart("bubbleChart");
            bool scatter = layer.IsChart("scatterChart");
            foreach (var series in layer.Children().ToList().Where(element => element.LocalName == "ser")) {
                int? seriesLength = null;
                OfficeXmlNode[] caches = bubble
                    ? new[] { series.Child(NsC, "xVal"), series.Child(NsC, "yVal"), series.Child(NsC, "bubbleSize") }
                    : scatter
                        ? new[] { series.Child(NsC, "xVal"), series.Child(NsC, "yVal") }
                        : new[] { series.Child(NsC, "val") };
                foreach (var cache in caches) {
                    if (cache.IsNone) return true;
                    var points = OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(CachePoints(cache, numeric: true), maximumPoints);
                    int length = OfficeOpenXmlChartCacheReader.GetCachedPointLength(cache, points, point => point.UIntAttr("idx"), maximumPoints);
                    // An empty cell leaves its plotted value out of the cache, which the chart's blank setting draws.
                    // Bubble values and scatter X positions have no such rule.
                    bool allowsEmptyCells = !bubble && (cache.IsChart("val") || cache.IsChart("yVal"));
                    if (length == 0 || points.Count != length && !allowsEmptyCells ||
                        points.Any(point => point.UIntAttr("idx") == null) ||
                        points.Select(point => point.UIntAttr("idx")!.Value).Distinct().Count() != points.Count ||
                        points.Any(point => !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(point.ChildText(NsC, "v"), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number) || double.IsNaN(number) || double.IsInfinity(number))) return true;
                    if (seriesLength.HasValue && seriesLength.Value != length) return true;
                    seriesLength = length;
                }
                if (!seriesLength.HasValue) return true;
                if (scatter || bubble) continue;
                if (!series.Child(NsC, "cat").Child(NsC, "multiLvlStrRef").IsNone) return true;
                var categories = OfficeOpenXmlChartCacheReader.ReadCachedStrings(series.Child(NsC, "cat"), maximumPoints);
                if (categories.Count != seriesLength.Value ||
                    sharedCategories != null && !sharedCategories.SequenceEqual(categories, StringComparer.Ordinal)) return true;
                sharedCategories = categories;
            }
            return false;
        }

        // SDK callers
        internal static Result? ReadPlot(ChartPart part, C.Chart chart, A.ColorScheme? scheme, int maximumPoints,
            out OfficeChartType kind, out double bubbleScale, out OfficeChartBubbleSizeMode bubbleMode) =>
            ReadPlot(new OfficeSdkChartPart(part), OfficeXmlNode.FromSdk(chart), OfficeXmlNode.FromSdk(scheme), maximumPoints,
                out kind, out bubbleScale, out bubbleMode);
    }
}
