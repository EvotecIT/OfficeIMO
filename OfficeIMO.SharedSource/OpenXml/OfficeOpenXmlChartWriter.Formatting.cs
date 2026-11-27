using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static void PreserveSharedChartFormatting(
            OfficeXmlNode source, OfficeXmlNode replacement,
            ISet<uint> preservedSeriesIndexes) {
            var axisBindings = new Dictionary<uint, uint>();
            PreserveSharedChartLayers(
                source, replacement, preservedSeriesIndexes, axisBindings);
            PreserveSharedAxes(source, replacement, axisBindings);
        }

        private static List<OfficeXmlNode> SharedChartLayers(OfficeXmlNode plotArea) {
            var layers = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in plotArea.Children()) {
                if (IsSharedChartLayer(child)) layers.Add(child);
            }
            return layers;
        }

        // Only applied to children of a shared chart layer, where every c:ser is one of the SDK's shared series classes.
        private static List<OfficeXmlNode> SharedSeriesOf(OfficeXmlNode layer) => layer.Children(NsC, "ser").ToList();

        private static void PreserveSharedChartLayers(
            OfficeXmlNode source, OfficeXmlNode replacement,
            ISet<uint> preservedSeriesIndexes, IDictionary<uint, uint> axisBindings) {
            List<OfficeXmlNode> sourceLayers = SharedChartLayers(source);
            var sourceLayerPositions = sourceLayers.Select((layer, index) => (layer, index))
                .ToDictionary(item => item.layer, item => item.index);
            var sourceGroups = OfficeOpenXmlChartAxisGroups.Create(source);
            var replacementGroups = OfficeOpenXmlChartAxisGroups.Create(replacement);
            var requestedSeries = SharedChartLayers(replacement)
                .SelectMany(layer => SharedSeriesOf(layer)
                    .Select(series => (LayerType: layer.LocalName, Index: series.Child(NsC, "idx").UIntVal(),
                        Group: replacementGroups.Read(layer))))
                .Where(item => item.Index.HasValue).ToList();
            var usedLayers = new HashSet<OfficeXmlNode>();
            var sourceOrder = new Dictionary<OfficeXmlNode, int>();
            foreach (OfficeXmlNode generated in SharedChartLayers(replacement)) {
                List<OfficeXmlNode> matches = sourceLayers.Where(candidate =>
                    !usedLayers.Contains(candidate) &&
                    AreCompatibleSharedChartLayers(candidate, generated, sourceGroups, replacementGroups)).ToList();
                if (matches.Count == 0) continue;
                List<OfficeXmlNode> generatedSeries = SharedSeriesOf(generated);
                int offset = 0;
                for (int layerIndex = 0; layerIndex < matches.Count && offset < generatedSeries.Count; layerIndex++) {
                    OfficeXmlNode match = matches[layerIndex];
                    int remaining = generatedSeries.Count - offset;
                    int oldCount = Math.Max(1, SharedSeriesOf(match).Count);
                    int reservedForLaterLayers = Math.Min(matches.Count - layerIndex - 1, remaining - 1);
                    int count = layerIndex == matches.Count - 1 ? remaining :
                        Math.Min(oldCount, remaining - reservedForLaterLayers);
                    // A shallow copy of the generated layer (it carries no attributes) holding this slice of its series.
                    OfficeXmlNode slice = replacement.CreateElement(NsC, generated.LocalName);
                    foreach (OfficeXmlAttribute attribute in generated.Attributes()) {
                        if (attribute.NamespaceUri.Length == 0) slice.SetAttribute(attribute.LocalName, attribute.Value);
                    }
                    foreach (OfficeXmlNode child in generated.Children()) {
                        if (!child.Is(NsC, "ser")) slice.Append(child.Clone());
                    }
                    foreach (OfficeXmlNode item in generatedSeries.Skip(offset).Take(count)) InsertSeries(slice, item.Clone());
                    OfficeXmlNode preserved = match.Clone();
                    ReplaceSharedSeriesData(preserved, slice, preservedSeriesIndexes);
                    BindSharedAxisReferences(match, generated, sourceGroups, replacementGroups, axisBindings);
                    ReplaceSharedAxisReferences(preserved, generated);
                    preserved = replacement.InsertBefore(preserved, generated);
                    sourceOrder.Add(preserved, sourceLayerPositions[match]);
                    usedLayers.Add(match);
                    offset += count;
                }
                generated.Remove();
            }
            // A targeted series in a native sibling layer must not silently lose its
            // distinct axis pair when the replacement has fewer axis groups.
            if (sourceLayers.Where(layer => !usedLayers.Contains(layer)).Any(layer =>
                SharedSeriesOf(layer)
                    .Any(series => requestedSeries.Any(requested => requested.LayerType == layer.LocalName &&
                        requested.Index == series.Child(NsC, "idx").UIntVal() &&
                        requested.Group != sourceGroups.Read(layer)))))
                throw new NotSupportedException("A native chart layer with a distinct axis pair cannot be flattened during a shared data update.");
            // Retained layers keep their source-relative overlap order. Keep unmatched
            // generated layers in their original slots so a newly added area layer can
            // still sit behind an existing line rather than being moved to the end.
            List<OfficeXmlNode> layers = SharedChartLayers(replacement);
            OfficeXmlNode[] retained = layers.Where(sourceOrder.ContainsKey)
                .OrderBy(layer => sourceOrder[layer]).ToArray();
            int retainedIndex = 0;
            var ordered = layers.Select(layer => sourceOrder.ContainsKey(layer)
                ? retained[retainedIndex++] : layer).ToList();
            if (layers.Count > 0) {
                OfficeXmlNode following = layers[layers.Count - 1].NextSibling;
                foreach (OfficeXmlNode layer in layers) layer.Remove();
                foreach (OfficeXmlNode layer in ordered) replacement.InsertBefore(layer, following);
            }
        }

        private static bool IsSharedChartLayer(OfficeXmlNode element) =>
            element.NamespaceId == NsC && element.LocalName is "barChart" or "lineChart" or "areaChart" or
                "radarChart" or "pieChart" or "doughnutChart" or "bubbleChart";

        private static bool AreCompatibleSharedChartLayers(OfficeXmlNode source,
            OfficeXmlNode replacement, OfficeOpenXmlChartAxisGroups.Groups sourceGroups,
            OfficeOpenXmlChartAxisGroups.Groups replacementGroups) {
            if (source.LocalName != replacement.LocalName) return false;
            switch (source.LocalName) {
                case "barChart":
                    if (source.ChildVal(NsC, "barDir") != replacement.ChildVal(NsC, "barDir") ||
                        source.ChildVal(NsC, "grouping") != replacement.ChildVal(NsC, "grouping")) return false;
                    break;
                case "lineChart":
                case "areaChart":
                    if (source.ChildVal(NsC, "grouping") != replacement.ChildVal(NsC, "grouping")) return false;
                    break;
            }

            return sourceGroups.Read(source) == replacementGroups.Read(replacement);
        }

        /// <summary>The child order of a shared chart layer.</summary>
        private static OfficeXmlChildOrder SharedLayerOrder(OfficeXmlNode layer) => layer.LocalName switch {
            "barChart" => CBarChartView.Order,
            "lineChart" => CLineChartView.Order,
            "areaChart" => CAreaChartView.Order,
            "radarChart" => CRadarChartView.Order,
            "pieChart" => CPieChartView.Order,
            "doughnutChart" => CDoughnutChartView.Order,
            "bubbleChart" => CBubbleChartView.Order,
            _ => throw new InvalidOperationException("The element '" + layer.LocalName + "' is not a shared chart layer.")
        };

        private static void ReplaceSharedSeriesData(
            OfficeXmlNode preserved,
            OfficeXmlNode generated,
            ISet<uint> preservedSeriesIndexes) {
            List<OfficeXmlNode> existingSeries = SharedSeriesOf(preserved);
            OfficeXmlNode insertionPoint = existingSeries.FirstOrDefault();
            List<OfficeXmlNode> oldSeries = existingSeries
                .OrderBy(series => series.Child(NsC, "order").UIntVal() ?? uint.MaxValue).ToList();
            List<OfficeXmlNode> generatedSeriesElements = SharedSeriesOf(generated);
            OfficeXmlChildOrder? seriesOrder = SeriesOrder(preserved.LocalName);
            for (int position = 0; position < generatedSeriesElements.Count; position++) {
                OfficeXmlNode generatedSeries =
                    generatedSeriesElements[position];
                uint? seriesIndex = generatedSeries.Child(NsC, "idx").UIntVal();
                // Both layers are the same kind, so their series are the same SDK series class.
                OfficeXmlNode sourceSeries = position < oldSeries.Count ? oldSeries[position] : default;
                OfficeXmlNode updated = sourceSeries.IsNone
                    ? generatedSeries.Clone()
                    : UpdateSharedSeriesData(sourceSeries, generatedSeries, seriesOrder);
                if (!sourceSeries.IsNone) {
                    if (seriesIndex.HasValue) {
                        preservedSeriesIndexes.Add(seriesIndex.Value);
                    }
                }
                // A layer without series takes each one in schema order after the ones already added.
                if (insertionPoint.IsNone) preserved.InsertOrdered(updated, SharedLayerOrder(preserved));
                else preserved.InsertBefore(updated, insertionPoint);
            }
            foreach (OfficeXmlNode series in oldSeries) series.Remove();
        }

        private static OfficeXmlNode UpdateSharedSeriesData(OfficeXmlNode source,
            OfficeXmlNode generated, OfficeXmlChildOrder? seriesOrder) {
            OfficeXmlNode updated = source.Clone();
            ReplaceSharedSeriesChild(updated, generated, "idx", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "order", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "tx", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "cat", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "val", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "xVal", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "yVal", seriesOrder);
            ReplaceSharedSeriesChild(updated, generated, "bubbleSize", seriesOrder);
            return updated;
        }

        private static void ReplaceSharedSeriesChild(OfficeXmlNode updated,
            OfficeXmlNode generated, string localName, OfficeXmlChildOrder? seriesOrder) {
            OfficeXmlNode current = updated.Child(NsC, localName);
            OfficeXmlNode replacement = generated.Child(NsC, localName);
            if (replacement.IsNone) {
                if (!current.IsNone) current.Remove();
            } else if (current.IsNone) {
                if (seriesOrder != null) SetOrdered(updated, replacement.Clone(), seriesOrder);
                else updated.Append(replacement.Clone());
            } else {
                current.ReplaceWith(replacement.Clone());
            }
        }

        private static void ReplaceSharedAxisReferences(OfficeXmlNode preserved,
            OfficeXmlNode generated) {
            List<OfficeXmlNode> preservedIds = preserved.Children(NsC, "axId").ToList();
            List<OfficeXmlNode> generatedIds = generated.Children(NsC, "axId").ToList();
            if (preservedIds.Count != generatedIds.Count) return;
            for (int index = 0; index < preservedIds.Count; index++) {
                preservedIds[index].SetAttribute("val", generatedIds[index].Val());
            }
        }

        private static void BindSharedAxisReferences(OfficeXmlNode source,
            OfficeXmlNode generated, OfficeOpenXmlChartAxisGroups.Groups sourceGroups,
            OfficeOpenXmlChartAxisGroups.Groups generatedGroups,
            IDictionary<uint, uint> bindings) {
            List<OfficeXmlNode> sourceIds = source.Children(NsC, "axId").ToList();
            List<OfficeXmlNode> generatedIds = generated.Children(NsC, "axId").ToList();
            if (sourceIds.Count != generatedIds.Count) return;
            for (int index = 0; index < sourceIds.Count; index++) {
                uint? generatedId = generatedIds[index].UIntVal();
                uint? sourceId = sourceIds[index].UIntVal();
                if (!source.Is(NsC, "bubbleChart") && !source.Is(NsC, "scatterChart")) {
                    OfficeXmlNode generatedAxis = generatedGroups.Resolve(generatedId);
                    if (generatedAxis.IsNone) throw new NotSupportedException("A chart axis reference must resolve to a unique category or value axis.");
                    var matches = sourceIds.Select(id => sourceGroups.Resolve(id.UIntVal())).Where(axis => !axis.IsNone &&
                        (generatedAxis.Is(NsC, "valAx") ? axis.Is(NsC, "valAx") : IsSharedCategoryAxis(axis))).Take(2).ToList();
                    if (matches.Count != 1) throw new NotSupportedException("A category chart layer must reference one category axis and one value axis.");
                    sourceId = matches[0].Child(NsC, "axId").UIntVal()!.Value;
                }
                if (!sourceId.HasValue || !generatedId.HasValue) continue;
                if (bindings.TryGetValue(generatedId.Value, out uint previous) && previous != sourceId.Value)
                    throw new NotSupportedException("Native layers in the same chart family and axis group must share their axis references for a shared data update.");
                bindings[generatedId.Value] = sourceId.Value;
            }
        }

        private static bool IsSharedCategoryAxis(OfficeXmlNode axis) =>
            axis.Is(NsC, "catAx") || axis.Is(NsC, "dateAx");

        private static bool IsSharedAxis(OfficeXmlNode axis) => axis.Is(NsC, "valAx") || IsSharedCategoryAxis(axis);

        private static bool UsesHorizontalSharedAxes(OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode chart in plotArea.Children(NsC, "barChart")) {
                if (chart.ChildVal(NsC, "barDir") == "bar") return true;
            }
            return false;
        }

        private static void PreserveSharedAxes(OfficeXmlNode source, OfficeXmlNode replacement,
            IReadOnlyDictionary<uint, uint> bindings) {
            if (UsesHorizontalSharedAxes(source) != UsesHorizontalSharedAxes(replacement)) return;
            List<OfficeXmlNode> sourceAxes = source.Children().ToList().Where(IsSharedAxis).ToList();
            List<OfficeXmlNode> generatedAxes = replacement.Children().ToList().Where(IsSharedAxis).ToList();
            var used = new HashSet<OfficeXmlNode>();
            foreach (OfficeXmlNode generated in generatedAxes) {
                bool Compatible(OfficeXmlNode axis) => !used.Contains(axis) &&
                    (axis.LocalName == generated.LocalName || IsSharedCategoryAxis(axis) && IsSharedCategoryAxis(generated));
                OfficeXmlNode match = default;
                uint? generatedId = generated.Child(NsC, "axId").UIntVal();
                if (generatedId.HasValue && bindings.TryGetValue(generatedId.Value, out uint sourceId)) {
                    match = sourceAxes.FirstOrDefault(axis => Compatible(axis) && axis.Child(NsC, "axId").UIntVal() == sourceId);
                } else {
                    // A family change can still retain an unambiguous axis in the same
                    // orientation. XML element order never determines axis identity.
                    var candidates = sourceAxes.Where(axis => Compatible(axis) &&
                        axis.ChildVal(NsC, "axPos") == generated.ChildVal(NsC, "axPos")).Take(2).ToList();
                    if (candidates.Count == 1) match = candidates[0];
                }
                if (match.IsNone) continue;
                used.Add(match);
                ReplaceSharedAxis(replacement, match, generated);
            }
        }
        private static void ReplaceSharedAxis(OfficeXmlNode replacement, OfficeXmlNode source,
            OfficeXmlNode generated) {
            OfficeXmlNode preserved = source.Clone();
            OfficeXmlNode generatedId = generated.Child(NsC, "axId");
            OfficeXmlNode generatedCrossing = generated.Child(NsC, "crossAx");
            if (!generatedId.IsNone) {
                OfficeXmlNode id = preserved.Child(NsC, "axId");
                if (!id.IsNone) id.Remove();
                preserved.InsertBefore(generatedId.Clone(), preserved.FirstChild);
            }
            if (!generatedCrossing.IsNone) {
                OfficeXmlNode crossing = preserved.Child(NsC, "crossAx");
                if (!crossing.IsNone) crossing.SetAttribute("val", generatedCrossing.Val());
            }
            generated.ReplaceWith(preserved);
        }
    }
}
