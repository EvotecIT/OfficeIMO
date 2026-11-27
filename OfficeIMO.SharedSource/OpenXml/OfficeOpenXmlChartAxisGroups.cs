using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Identifies category axis groups from the referenced plot structure.</summary>
    internal static class OfficeOpenXmlChartAxisGroups {
        internal static OfficeChartAxisGroup Read(OfficeXmlNode plotArea, OfficeXmlNode layer) {
            return Create(plotArea).Read(layer);
        }

        /// <summary>Captures axis identity once for one read or update operation.</summary>
        internal static Groups Create(OfficeXmlNode plotArea) {
            var allReferences = new HashSet<uint>();
            foreach (OfficeXmlNode element in plotArea.Children()) {
                if (!element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase)) continue;
                foreach (OfficeXmlNode axis in element.Children(NsC, "axId")) {
                    if (axis.UIntVal() is uint id) allReferences.Add(id);
                }
            }
            var valueAxes = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode axis in plotArea.Children(NsC, "valAx")) {
                if (axis.Child(NsC, "axId").UIntVal() is uint id && allReferences.Contains(id)) valueAxes.Add(axis);
            }
            var valueIds = new HashSet<uint>(valueAxes.Select(axis => axis.Child(NsC, "axId").UIntVal()!.Value));
            var secondary = new HashSet<uint>();
            if (valueIds.Count > 1 &&
                (!plotArea.Child(NsC, "catAx").IsNone || !plotArea.Child(NsC, "dateAx").IsNone)) {
                // The first category layer owns the primary pair. Axis XML order
                // and displayed sides can both change independently.
                uint[] pairedValues = plotArea.Children().ToList()
                    .Where(axis => axis.IsChart("catAx") || axis.IsChart("dateAx"))
                    .Select(axis => axis.Child(NsC, "crossAx").UIntVal())
                    .Where(id => id.HasValue && valueIds.Contains(id.Value))
                    .Select(id => id!.Value).Distinct().ToArray();
                if (pairedValues.Length != valueIds.Count || pairedValues.Length != 2)
                    throw new NotSupportedException("The chart axis pairs cannot be classified safely.");
                OfficeXmlNode firstLayer = default;
                foreach (OfficeXmlNode axis in plotArea.Children()) {
                    if (axis.IsChart("barChart") || axis.IsChart("lineChart") || axis.IsChart("areaChart") || axis.IsChart("radarChart")) {
                        firstLayer = axis;
                        break;
                    }
                }
                uint[] firstValues = firstLayer.IsNone ? Array.Empty<uint>() : firstLayer.Children(NsC, "axId").ToList()
                    .Select(axis => axis.UIntVal())
                    .Where(id => id.HasValue && valueIds.Contains(id.Value))
                    .Select(id => id!.Value).Distinct().ToArray();
                if (firstValues.Length != 1)
                    throw new NotSupportedException("The primary chart axis pair cannot be classified safely.");
                secondary.Add(pairedValues.Single(id => id != firstValues[0]));
            }
            var axes = plotArea.Children().ToList()
                .Where(axis => (axis.IsChart("valAx") || axis.IsChart("catAx") || axis.IsChart("dateAx")) &&
                    axis.Child(NsC, "axId").UIntVal() != null)
                .GroupBy(axis => axis.Child(NsC, "axId").UIntVal()!.Value)
                .ToDictionary(group => group.Key, group => group.ToArray());
            return new Groups(secondary, axes);
        }

        internal sealed class Groups {
            private readonly HashSet<uint> _secondary;
            private readonly IReadOnlyDictionary<uint, OfficeXmlNode[]> _axes;
            internal Groups(HashSet<uint> secondary, IReadOnlyDictionary<uint, OfficeXmlNode[]> axes) {
                _secondary = secondary; _axes = axes;
            }
            internal OfficeXmlNode Resolve(uint? id) {
                if (!id.HasValue || !_axes.TryGetValue(id.Value, out OfficeXmlNode[]? axes) || axes == null) return default;
                if (axes.Length != 1) throw new NotSupportedException("A chart axis reference must resolve to a unique axis.");
                return axes[0];
            }
            internal OfficeChartAxisGroup Read(OfficeXmlNode layer) {
                if (layer.IsChart("scatterChart") || layer.IsChart("bubbleChart")) return OfficeChartAxisGroup.Primary;
                foreach (OfficeXmlNode axis in layer.Children(NsC, "axId")) {
                    if (axis.UIntVal() is uint id && _secondary.Contains(id)) return OfficeChartAxisGroup.Secondary;
                }
                return OfficeChartAxisGroup.Primary;
            }
        }

        // SDK callers
        internal static OfficeChartAxisGroup Read(C.PlotArea plotArea, OpenXmlCompositeElement layer) =>
            Read(OfficeXmlNode.FromSdk(plotArea), OfficeXmlNode.FromSdk(layer));

        internal static SdkGroups Create(C.PlotArea plotArea) => new(Create(OfficeXmlNode.FromSdk(plotArea)));

        /// <summary>The axis groups of an SDK plot area, resolving to SDK elements for callers that still hold them.</summary>
        internal sealed class SdkGroups {
            private readonly Groups _groups;
            internal SdkGroups(Groups groups) => _groups = groups;
            internal OpenXmlCompositeElement? Resolve(uint? id) => (OpenXmlCompositeElement?)_groups.Resolve(id).SdkElement;
            internal OfficeChartAxisGroup Read(OpenXmlCompositeElement layer) => _groups.Read(OfficeXmlNode.FromSdk(layer));
        }
    }
}
