using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static void UpdateScatterChartLayers(OfficeXmlNode plotArea,
            IReadOnlyList<OfficeXmlNode> scatterCharts, OfficeChartData data) {
            int seriesOffset = 0;
            for (int layerIndex = 0; layerIndex < scatterCharts.Count; layerIndex++) {
                OfficeXmlNode scatterChart = scatterCharts[layerIndex];
                int remainingSeries = data.Series.Count - seriesOffset;
                if (remainingSeries <= 0) {
                    scatterChart.Remove();
                    continue;
                }

                int currentLayerSize = Math.Max(1, scatterChart.Children(NsC, "ser").Count);
                int reservedForLaterLayers = Math.Min(scatterCharts.Count - layerIndex - 1, remainingSeries - 1);
                int seriesCount = layerIndex == scatterCharts.Count - 1
                    ? remainingSeries
                    : Math.Min(currentLayerSize, remainingSeries - reservedForLaterLayers);
                OfficeXmlNode style = scatterChart.Child(NsC, "scatterStyle");
                bool created = style.IsNone;
                if (created) style = scatterChart.CreateElement(NsC, "scatterStyle");
                // The SDK reads an unknown token as an error, as it did before.
                string? current = created ? null : ReadScatterStyle(style.Val());
                style.SetAttribute("val", current is "smooth" or "smoothMarker" ? "smoothMarker" : "lineMarker");
                if (created) SetOrdered(scatterChart, style, CScatterChartView.Order);
                UpdateScatterChartSeries(scatterChart, data, seriesOffset, seriesCount);
                seriesOffset += seriesCount;
            }

            RemoveUnusedScatterAxes(plotArea);
        }

        // c:scatterStyle tokens; any other value makes the SDK throw FormatException when it is read.
        private static string? ReadScatterStyle(string? token) {
            if (token == null) return null;
            return token is "none" or "line" or "lineMarker" or "marker" or "smooth" or "smoothMarker"
                ? token
                : throw new FormatException("The scatter style '" + token + "' is not a valid value.");
        }

        private static void UpdateScatterChartSeries(OfficeXmlNode scatterChart,
            OfficeChartData data, int seriesOffset, int seriesCount) {
            List<OfficeXmlNode> existingSeries = scatterChart.Children(NsC, "ser").ToList();
            OfficeXmlNode template = existingSeries.Count > 0 ? existingSeries[existingSeries.Count - 1] : default;

            for (int localIndex = 0; localIndex < seriesCount; localIndex++) {
                OfficeXmlNode seriesElement;
                if (localIndex < existingSeries.Count) {
                    seriesElement = existingSeries[localIndex];
                } else {
                    seriesElement = !template.IsNone
                        ? template.Clone()
                        : scatterChart.CreateElement(NsC, "ser");
                    ClearAddedSeriesMetadata(seriesElement);
                    seriesElement = InsertSeries(scatterChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                int seriesIndex = seriesOffset + localIndex;
                UpdateSeriesIndexOrder(seriesElement, seriesIndex);
                UpdateScatterSeriesText(seriesElement, seriesIndex, data.Series[seriesIndex].Name);
                UpdateXValues(seriesElement, seriesIndex, data.Series[seriesIndex].XValues!);
                UpdateYValues(seriesElement, seriesIndex, data.Series[seriesIndex].Values);
            }

            for (int localIndex = existingSeries.Count - 1; localIndex >= seriesCount; localIndex--) {
                existingSeries[localIndex].Remove();
            }
        }

        private static void RemoveUnusedScatterAxes(OfficeXmlNode plotArea) {
            var usedAxisIds = new HashSet<uint>();
            foreach (OfficeXmlNode chart in plotArea.Children(NsC, "scatterChart")) {
                foreach (OfficeXmlNode axis in chart.Children(NsC, "axId")) {
                    uint? id = axis.UIntVal();
                    if (id.HasValue) usedAxisIds.Add(id.Value);
                }
            }
            foreach (OfficeXmlNode axis in plotArea.Children(NsC, "valAx").ToList()) {
                uint? axisId = axis.Child(NsC, "axId").UIntVal();
                if (axisId.HasValue && !usedAxisIds.Contains(axisId.Value)) axis.Remove();
            }
        }
    }
}
