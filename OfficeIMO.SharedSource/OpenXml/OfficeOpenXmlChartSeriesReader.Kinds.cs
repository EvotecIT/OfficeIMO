using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        /// <summary>Reads the shared data family; 3-D groups still require separate projection qualification.</summary>
        internal static bool TryReadKind(OpenXmlElement group, out OfficeChartKind kind) {
            kind = default;
            switch (group.LocalName) {
                case "barChart":
                case "bar3DChart":
                    bool horizontal = group.GetFirstChild<C.BarDirection>()?.Val?.Value == C.BarDirectionValues.Bar;
                    var barGrouping = group.GetFirstChild<C.BarGrouping>()?.Val?.Value;
                    kind = barGrouping == C.BarGroupingValues.Stacked
                        ? horizontal ? OfficeChartKind.BarStacked : OfficeChartKind.ColumnStacked
                        : barGrouping == C.BarGroupingValues.PercentStacked
                            ? horizontal ? OfficeChartKind.BarStacked100 : OfficeChartKind.ColumnStacked100
                            : horizontal ? OfficeChartKind.BarClustered : OfficeChartKind.ColumnClustered;
                    return true;
                case "lineChart":
                case "line3DChart":
                case "areaChart":
                case "area3DChart":
                    bool area = group.LocalName.StartsWith("area", System.StringComparison.Ordinal);
                    var grouping = group.GetFirstChild<C.Grouping>()?.Val?.Value;
                    kind = grouping == C.GroupingValues.Stacked
                        ? area ? OfficeChartKind.AreaStacked : OfficeChartKind.LineStacked
                        : grouping == C.GroupingValues.PercentStacked
                            ? area ? OfficeChartKind.AreaStacked100 : OfficeChartKind.LineStacked100
                            : area ? OfficeChartKind.Area : OfficeChartKind.Line;
                    return true;
                case "radarChart": kind = OfficeChartKind.Radar; return true;
                case "scatterChart": kind = OfficeChartKind.Scatter; return true;
                case "pieChart":
                case "pie3DChart": kind = OfficeChartKind.Pie; return true;
                case "doughnutChart": kind = OfficeChartKind.Doughnut; return true;
                case "bubbleChart": kind = OfficeChartKind.Bubble; return true;
                default: return false;
            }
        }
    }
}
