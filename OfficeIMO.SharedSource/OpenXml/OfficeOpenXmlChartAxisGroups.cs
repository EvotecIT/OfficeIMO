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
            if (layer is C.ScatterChart || layer is C.BubbleChart) return OfficeChartAxisGroup.Primary;
            var allReferences = new HashSet<uint>(plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase))
                .SelectMany(element => element.Elements<C.AxisId>()).Where(axis => axis.Val != null).Select(axis => axis.Val!.Value));
            var valueAxes = plotArea.Elements<C.ValueAxis>()
                .Where(axis => axis.AxisId?.Val != null && allReferences.Contains(axis.AxisId.Val.Value)).ToList();
            if (valueAxes.Select(axis => axis.AxisId!.Val!.Value).Distinct().Count() <= 1) return OfficeChartAxisGroup.Primary;
            // The category/value pairs, rather than their displayed sides, identify
            // the primary and secondary groups. A primary value axis may be moved.
            var valueIds = new HashSet<uint>(valueAxes.Select(axis => axis.AxisId!.Val!.Value));
            uint[] orderedPairs = plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(axis => axis is C.CategoryAxis or C.DateAxis)
                .Select(axis => axis.GetFirstChild<C.CrossingAxis>()?.Val?.Value)
                .Where(id => id.HasValue && valueIds.Contains(id.Value))
                .Select(id => id!.Value).Distinct().ToArray();
            if (orderedPairs.Length != valueIds.Count || orderedPairs.Length > 2)
                throw new NotSupportedException("The chart axis pairs cannot be classified safely.");
            var references = new HashSet<uint>(layer.Elements<C.AxisId>().Where(axis => axis.Val != null).Select(axis => axis.Val!.Value));
            if (references.Contains(orderedPairs[0]) == references.Contains(orderedPairs[1]))
                throw new NotSupportedException("A chart layer must reference one value-axis group.");
            return references.Contains(orderedPairs[1])
                ? OfficeChartAxisGroup.Secondary : OfficeChartAxisGroup.Primary;
        }
    }
}
