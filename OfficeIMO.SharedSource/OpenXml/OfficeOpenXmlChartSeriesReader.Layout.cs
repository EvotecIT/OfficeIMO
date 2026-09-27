using System;
using System.Linq;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static OfficeChartLayout ReadLayout(C.Chart chart, OfficeChartKind kind) {
            var legend = chart.GetFirstChild<C.Legend>();
            var position = legend?.GetFirstChild<C.LegendPosition>()?.Val?.Value;
            var sharedPosition = position == C.LegendPositionValues.Left ? OfficeChartLegendPosition.Left :
                position == C.LegendPositionValues.Bottom ? OfficeChartLegendPosition.Bottom :
                position == C.LegendPositionValues.Top ? OfficeChartLegendPosition.Top : OfficeChartLegendPosition.Right;
            bool radial = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut;
            var hidden = radial ? legend?.Elements<C.LegendEntry>().Where(entry => entry.GetFirstChild<C.Delete>() is C.Delete delete && delete.Val?.Value != false)
                .Select(entry => entry.GetFirstChild<C.Index>()?.Val?.Value).Where(value => value.HasValue && value.Value <= int.MaxValue)
                .Select(value => (int)value!.Value).ToArray() : null;
            return new OfficeChartLayout(overlayLegend: legend?.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false,
                overlayTitle: chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Overlay>() is C.Overlay title && title.Val?.Value != false,
                showLegend: legend != null, legendPosition: sharedPosition, hiddenCategoryLegendIndexes: hidden);
        }
    }
}
