using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native radial geometry codec for the Office format libraries.</summary>
internal static class OfficeOpenXmlChartRadialLayout {
    internal static OfficeChartRadialLayout Read(C.Chart? chart) {
        C.PlotArea? plot = chart?.PlotArea;
        C.DoughnutChart? doughnut = plot?.GetFirstChild<C.DoughnutChart>();
        C.PieChart? pie = plot?.GetFirstChild<C.PieChart>();
        int angle = doughnut?.GetFirstChild<C.FirstSliceAngle>()?.Val?.Value ??
            pie?.GetFirstChild<C.FirstSliceAngle>()?.Val?.Value ?? 0;
        int hole = doughnut?.GetFirstChild<C.HoleSize>()?.Val?.Value ?? 50;
        return new OfficeChartRadialLayout(angle, hole);
    }

    internal static void Apply(C.Chart? chart, OfficeChartRadialLayout layout) {
        if (layout == null) throw new ArgumentNullException(nameof(layout));
        C.PlotArea plot = chart?.PlotArea ?? throw new InvalidOperationException("The chart has no plot area.");
        var layers = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList();
        if (layers.Count == 0 || layers.Any(element => element is not C.PieChart && element is not C.DoughnutChart))
            throw new NotSupportedException("Radial geometry requires a native two-dimensional pie or doughnut chart.");
        foreach (var layer in layers) {
            layer.GetFirstChild<C.FirstSliceAngle>()?.Remove();
            layer.AddChild(new C.FirstSliceAngle { Val = (ushort)layout.FirstSliceAngleDegrees }, true);
            if (layer is C.DoughnutChart) {
                layer.GetFirstChild<C.HoleSize>()?.Remove();
                layer.AddChild(new C.HoleSize { Val = (byte)layout.DoughnutHolePercent }, true);
            }
        }
    }
}
