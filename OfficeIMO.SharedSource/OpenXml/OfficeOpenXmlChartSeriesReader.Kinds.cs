using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        /// <summary>Describes what a flat projection of a native chart group leaves out, or null when nothing is.</summary>
        internal static string? DescribeProjection(OfficeXmlNode group) {
            string? omitted = group.LocalName switch {
                "bar3DChart" or "line3DChart" or "area3DChart" or "pie3DChart" =>
                    "3-D depth, perspective and lighting",
                "ofPieChart" => "the secondary pie or bar that details the split slices",
                "surfaceChart" or "surface3DChart" => "surface bands and their colour scale; the series are drawn as lines",
                "stockChart" when !group.Child(NsC, "upDownBars").IsNone => "up and down bars",
                _ => null
            };
            return omitted == null ? null :
                "The native " + group.LocalName + " is drawn as a flat 2-D chart without " + omitted + ".";
        }

        /// <summary>Reads the shared data family; 3-D groups still require separate projection qualification.</summary>
        internal static bool TryReadKind(OfficeXmlNode group, out OfficeChartType kind) {
            kind = default;
            switch (group.LocalName) {
                case "barChart":
                case "bar3DChart":
                    bool horizontal = group.ChildVal(NsC, "barDir") == "bar";
                    string? barGrouping = group.ChildVal(NsC, "grouping");
                    kind = barGrouping == "stacked"
                        ? horizontal ? OfficeChartType.BarStacked : OfficeChartType.ColumnStacked
                        : barGrouping == "percentStacked"
                            ? horizontal ? OfficeChartType.BarStacked100 : OfficeChartType.ColumnStacked100
                            : horizontal ? OfficeChartType.BarClustered : OfficeChartType.ColumnClustered;
                    return true;
                case "lineChart":
                case "line3DChart":
                case "areaChart":
                case "area3DChart":
                    bool area = group.LocalName.StartsWith("area", System.StringComparison.Ordinal);
                    string? grouping = group.ChildVal(NsC, "grouping");
                    kind = grouping == "stacked"
                        ? area ? OfficeChartType.AreaStacked : OfficeChartType.LineStacked
                        : grouping == "percentStacked"
                            ? area ? OfficeChartType.AreaStacked100 : OfficeChartType.LineStacked100
                            : area ? OfficeChartType.Area : OfficeChartType.Line;
                    return true;
                // A stock chart's high-low-close series project onto a line chart with high-low lines.
                case "stockChart": kind = OfficeChartType.Line; return true;
                case "radarChart": kind = OfficeChartType.Radar; return true;
                case "scatterChart": kind = OfficeChartType.Scatter; return true;
                case "pieChart":
                case "pie3DChart": kind = OfficeChartType.Pie; return true;
                case "doughnutChart": kind = OfficeChartType.Doughnut; return true;
                case "bubbleChart": kind = OfficeChartType.Bubble; return true;
                default: return false;
            }
        }

        // SDK callers
        internal static string? DescribeProjection(OpenXmlElement group) => DescribeProjection(OfficeXmlNode.FromSdk(group));

        internal static bool TryReadKind(OpenXmlElement group, out OfficeChartType kind) => TryReadKind(OfficeXmlNode.FromSdk(group), out kind);
    }
}
