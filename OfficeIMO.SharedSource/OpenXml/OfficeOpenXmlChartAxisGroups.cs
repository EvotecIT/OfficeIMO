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
            var references = new HashSet<uint>(layer.Elements<C.AxisId>().Where(axis => axis.Val != null).Select(axis => axis.Val!.Value));
            return valueAxes.Any(axis => references.Contains(axis.AxisId!.Val!.Value) &&
                (axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Right || axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Top))
                ? OfficeChartAxisGroup.Secondary : OfficeChartAxisGroup.Primary;
        }
    }
}
