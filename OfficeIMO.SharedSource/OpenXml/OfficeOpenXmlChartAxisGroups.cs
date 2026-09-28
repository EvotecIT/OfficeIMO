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
            // The first category chart layer owns the primary pair. Axis XML order
            // and displayed sides can both change without changing layer identity.
            var valueIds = new HashSet<uint>(valueAxes.Select(axis => axis.AxisId!.Val!.Value));
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
            var references = new HashSet<uint>(layer.Elements<C.AxisId>().Where(axis => axis.Val != null).Select(axis => axis.Val!.Value));
            uint secondaryValue = pairedValues.Single(id => id != firstValues[0]);
            if (references.Contains(firstValues[0]) == references.Contains(secondaryValue))
                throw new NotSupportedException("A chart layer must reference one value-axis group.");
            return references.Contains(secondaryValue)
                ? OfficeChartAxisGroup.Secondary : OfficeChartAxisGroup.Primary;
        }
    }
}
