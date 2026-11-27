using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native radial geometry codec for the Office format libraries.</summary>
internal static class OfficeOpenXmlChartRadialLayout {
    internal static OfficeChartRadialLayout Read(OfficeXmlNode chart) {
        OfficeXmlNode plot = chart.Child(NsC, "plotArea");
        OfficeXmlNode doughnut = plot.Child(NsC, "doughnutChart");
        OfficeXmlNode pie = plot.Child(NsC, "pieChart");
        int angle = doughnut.Child(NsC, "firstSliceAng").UShortVal() ??
            pie.Child(NsC, "firstSliceAng").UShortVal() ?? 0;
        int hole = doughnut.Child(NsC, "holeSize").ByteVal() ?? 50;
        return new OfficeChartRadialLayout(angle, hole);
    }

    internal static void Apply(OfficeXmlNode chart, OfficeChartRadialLayout layout) {
        ArgumentNullException.ThrowIfNull(layout);
        OfficeXmlNode plot = chart.Child(NsC, "plotArea");
        if (plot.IsNone) throw new InvalidOperationException("The chart has no plot area.");
        var layers = plot.Children().ToList()
            .Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList();
        if (layers.Count == 0 || layers.Any(element => !element.IsChart("pieChart") && !element.IsChart("doughnutChart")))
            throw new NotSupportedException("Radial geometry requires a native two-dimensional pie or doughnut chart.");
        foreach (var layer in layers) {
            OfficeXmlChildOrder order = layer.IsChart("doughnutChart") ? CDoughnutChartView.Order : CPieChartView.Order;
            OfficeXmlNode existing = layer.Child(NsC, "firstSliceAng");
            if (!existing.IsNone) existing.Remove();
            SetOrdered(layer, layer.CreateVal(NsC, "firstSliceAng",
                ((ushort)layout.FirstSliceAngleDegrees).ToString(CultureInfo.InvariantCulture)), order);
            if (layer.IsChart("doughnutChart")) {
                OfficeXmlNode hole = layer.Child(NsC, "holeSize");
                if (!hole.IsNone) hole.Remove();
                SetOrdered(layer, layer.CreateVal(NsC, "holeSize",
                    ((byte)layout.DoughnutHolePercent).ToString(CultureInfo.InvariantCulture)), order);
            }
        }
    }

    // SDK callers
    internal static OfficeChartRadialLayout Read(C.Chart? chart) => Read(OfficeXmlNode.FromSdk(chart));

    internal static void Apply(C.Chart? chart, OfficeChartRadialLayout layout) => Apply(OfficeXmlNode.FromSdk(chart), layout);
}
