using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native bar spacing and category point placement into the shared chart layout.</summary>
internal static class OfficeOpenXmlChartCategoryGeometry {
    /// <summary>Applies the first bar group's gap width and overlap and the category crossing of its value axis.</summary>
    internal static OfficeChartLayout WithCategoryGeometry(this OfficeChartLayout layout, OfficeXmlNode plot,
        OfficeXmlNode scheme = default) {
        if (plot.IsNone) return layout;
        OfficeXmlNode highLow = default;
        foreach (OfficeXmlNode element in plot.Children()) {
            if (!(element.IsChart("lineChart") || element.IsChart("line3DChart") || element.IsChart("stockChart"))) continue;
            highLow = element.Child(NsC, "hiLowLines");
            if (!highLow.IsNone) break;
        }
        if (!highLow.IsNone) {
            OfficeXmlNode outline = highLow.Child(NsC, "spPr").Child(NsA, "ln");
            bool hidden = !outline.Child(NsA, "noFill").IsNone;
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), scheme);
            double? width = outline.IntAttr("w") is int emus && emus > 0 && emus <= 1584 * 12700 ? emus / 12700D : null;
            layout = layout.WithHighLowLines(!hidden, color, width);
        }
        OfficeXmlNode bars = default;
        foreach (OfficeXmlNode element in plot.Children()) {
            if (element.IsChart("barChart") || element.IsChart("bar3DChart")) {
                bars = element;
                break;
            }
        }
        if (!bars.IsNone) {
            int gap = bars.Child(NsC, "gapWidth").UShortVal() is ushort gapValue ? Math.Min((int)gapValue, 500) : 150;
            bool clustered = (bars.ChildVal(NsC, "grouping") ?? "clustered") == "clustered";
            // Stacked groups share one slot per stack, so their native 100% overlap has no further effect.
            int overlap = clustered && bars.Child(NsC, "overlap").SByteVal() is sbyte overlapValue
                ? Math.Max(-100, Math.Min(100, (int)overlapValue)) : 0;
            layout = layout.WithBarGeometry(gap, overlap);
        }
        OfficeXmlNode categoryLayer = default;
        foreach (OfficeXmlNode element in plot.Children()) {
            if (element.IsChart("lineChart") || element.IsChart("line3DChart") || element.IsChart("areaChart") ||
                element.IsChart("area3DChart") || element.IsChart("barChart") || element.IsChart("bar3DChart") ||
                element.IsChart("stockChart")) {
                categoryLayer = element;
                break;
            }
        }
        var axisIds = new List<uint>();
        if (!categoryLayer.IsNone) {
            foreach (OfficeXmlNode id in categoryLayer.Children(NsC, "axId")) axisIds.Add(id.UIntVal() ?? 0U);
        }
        OfficeXmlNode valueAxis = default;
        foreach (OfficeXmlNode axis in plot.Children(NsC, "valAx")) {
            if (axis.Child(NsC, "axId").UIntVal() is uint id && axisIds.Contains(id)) {
                valueAxis = axis;
                break;
            }
        }
        return layout.WithCategoryPointsOnTickMarks(valueAxis.ChildVal(NsC, "crossBetween") == "midCat");
    }

    // SDK callers
    internal static OfficeChartLayout WithCategoryGeometry(this OfficeChartLayout layout, C.PlotArea? plot,
        DocumentFormat.OpenXml.Drawing.ColorScheme? scheme = null) =>
        layout.WithCategoryGeometry(OfficeXmlNode.FromSdk(plot), OfficeXmlNode.FromSdk(scheme));
}
