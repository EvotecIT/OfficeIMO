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
            var secondary = valueAxes.Select(axis => axis.AxisId!.Val!.Value).Distinct().Count() <= 1
                ? new HashSet<uint>() : new HashSet<uint>(valueAxes.Where(axis =>
                    axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Right || axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Top)
                    .Select(axis => axis.AxisId!.Val!.Value));
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
