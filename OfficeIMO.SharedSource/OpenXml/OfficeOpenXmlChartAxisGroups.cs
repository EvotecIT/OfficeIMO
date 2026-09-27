using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Identifies category axis groups from the referenced plot structure.</summary>
    internal static class OfficeOpenXmlChartAxisGroups {
        internal static OfficeChartAxisGroup Read(C.PlotArea plotArea, OpenXmlCompositeElement layer) {
            return Create(plotArea).Read(layer);
        }

        /// <summary>Captures axis identity once for one read or update operation.</summary>
        internal static Groups Create(C.PlotArea plotArea) {
            var allReferences = new HashSet<uint>(plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase))
                .SelectMany(element => element.Elements<C.AxisId>()).Where(axis => axis.Val != null).Select(axis => axis.Val!.Value));
            var valueAxes = plotArea.Elements<C.ValueAxis>()
                .Where(axis => axis.AxisId?.Val != null && allReferences.Contains(axis.AxisId.Val.Value)).ToList();
            var valueIds = new HashSet<uint>(valueAxes.Select(axis => axis.AxisId!.Val!.Value));
            var secondary = new HashSet<uint>();
            if (valueIds.Count > 1 &&
                (plotArea.Elements<C.CategoryAxis>().Any() || plotArea.Elements<C.DateAxis>().Any())) {
                // The first category layer owns the primary pair. Axis XML order
                // and displayed sides can both change independently.
                uint[] pairedValues = plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                    .Where(axis => axis is C.CategoryAxis or C.DateAxis)
                    .Select(axis => axis.GetFirstChild<C.CrossingAxis>()?.Val?.Value)
                    .Where(id => id.HasValue && valueIds.Contains(id.Value))
                    .Select(id => id!.Value).Distinct().ToArray();
                if (pairedValues.Length != valueIds.Count || pairedValues.Length != 2)
                    throw new NotSupportedException("The chart axis pairs cannot be classified safely.");
                OpenXmlCompositeElement? firstLayer = plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                    .FirstOrDefault(axis => axis is C.BarChart or C.LineChart or C.AreaChart or C.RadarChart);
                uint[] firstValues = firstLayer?.Elements<C.AxisId>()
                    .Where(axis => axis.Val != null && valueIds.Contains(axis.Val.Value))
                    .Select(axis => axis.Val!.Value).Distinct().ToArray() ?? Array.Empty<uint>();
                if (firstValues.Length != 1)
                    throw new NotSupportedException("The primary chart axis pair cannot be classified safely.");
                secondary.Add(pairedValues.Single(id => id != firstValues[0]));
            }
            var axes = plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(axis => (axis is C.ValueAxis || axis is C.CategoryAxis || axis is C.DateAxis) &&
                    axis.GetFirstChild<C.AxisId>()?.Val != null)
                .GroupBy(axis => axis.GetFirstChild<C.AxisId>()!.Val!.Value)
                .ToDictionary(group => group.Key, group => group.ToArray());
            return new Groups(secondary, axes);
        }

        internal sealed class Groups {
            private readonly HashSet<uint> _secondary;
            private readonly IReadOnlyDictionary<uint, OpenXmlCompositeElement[]> _axes;
            internal Groups(HashSet<uint> secondary, IReadOnlyDictionary<uint, OpenXmlCompositeElement[]> axes) {
                _secondary = secondary; _axes = axes;
            }
            internal OpenXmlCompositeElement? Resolve(uint? id) {
                if (!id.HasValue || !_axes.TryGetValue(id.Value, out OpenXmlCompositeElement[]? axes) || axes == null) return null;
                if (axes.Length != 1) throw new NotSupportedException("A chart axis reference must resolve to a unique axis.");
                return axes[0];
            }
            internal OfficeChartAxisGroup Read(OpenXmlCompositeElement layer) =>
                layer is C.ScatterChart || layer is C.BubbleChart ||
                !layer.Elements<C.AxisId>().Any(axis => axis.Val != null && _secondary.Contains(axis.Val.Value))
                    ? OfficeChartAxisGroup.Primary : OfficeChartAxisGroup.Secondary;
        }
    }
}
