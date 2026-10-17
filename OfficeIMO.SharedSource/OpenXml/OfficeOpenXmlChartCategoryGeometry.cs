using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native bar spacing and category point placement into the shared chart layout.</summary>
internal static class OfficeOpenXmlChartCategoryGeometry {
    /// <summary>Applies the first bar group's gap width and overlap and the category crossing of its value axis.</summary>
    internal static OfficeChartLayout WithCategoryGeometry(this OfficeChartLayout layout, C.PlotArea? plot,
        DocumentFormat.OpenXml.Drawing.ColorScheme? scheme = null) {
        if (plot == null) return layout;
        C.HighLowLines? highLow = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(element => element is C.LineChart or C.Line3DChart or C.StockChart)
            .Select(element => element.GetFirstChild<C.HighLowLines>()).FirstOrDefault(lines => lines != null);
        if (highLow != null) {
            DocumentFormat.OpenXml.Drawing.Outline? outline = highLow.GetFirstChild<C.ChartShapeProperties>()?
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
            bool hidden = outline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.NoFill>() != null;
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                outline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>(), scheme);
            double? width = outline?.Width?.Value is int emus && emus > 0 && emus <= 1584 * 12700 ? emus / 12700D : null;
            layout = layout.WithHighLowLines(!hidden, color, width);
        }
        OpenXmlCompositeElement? bars = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .FirstOrDefault(element => element is C.BarChart or C.Bar3DChart);
        if (bars != null) {
            int gap = bars.GetFirstChild<C.GapWidth>()?.Val?.Value is ushort gapValue ? Math.Min((int)gapValue, 500) : 150;
            bool clustered = (bars.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered) ==
                C.BarGroupingValues.Clustered;
            // Stacked groups share one slot per stack, so their native 100% overlap has no further effect.
            int overlap = clustered && bars.GetFirstChild<C.Overlap>()?.Val?.Value is sbyte overlapValue
                ? Math.Max(-100, Math.Min(100, (int)overlapValue)) : 0;
            layout = layout.WithBarGeometry(gap, overlap);
        }
        OpenXmlCompositeElement? categoryLayer = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .FirstOrDefault(element => element is C.LineChart or C.Line3DChart or C.AreaChart or C.Area3DChart or
                C.BarChart or C.Bar3DChart or C.StockChart);
        uint[] axisIds = categoryLayer?.Elements<C.AxisId>().Select(id => id.Val?.Value ?? 0U).ToArray() ?? Array.Empty<uint>();
        C.ValueAxis? valueAxis = plot.Elements<C.ValueAxis>().FirstOrDefault(axis =>
            axis.AxisId?.Val?.Value is uint id && axisIds.Contains(id));
        return layout.WithCategoryPointsOnTickMarks(
            valueAxis?.GetFirstChild<C.CrossBetween>()?.Val?.Value == C.CrossBetweenValues.MidpointCategory);
    }
}
