using System;
using System.Collections.Generic;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        private static void UpdateScatterChartLayers(C.PlotArea plotArea,
            IReadOnlyList<C.ScatterChart> scatterCharts, PowerPointScatterChartData data) {
            int seriesOffset = 0;
            for (int layerIndex = 0; layerIndex < scatterCharts.Count; layerIndex++) {
                C.ScatterChart scatterChart = scatterCharts[layerIndex];
                int remainingSeries = data.Series.Count - seriesOffset;
                if (remainingSeries <= 0) {
                    scatterChart.Remove();
                    continue;
                }

                int futureLayers = scatterCharts.Count - layerIndex - 1;
                int seriesReservedForFutureLayers = Math.Min(futureLayers, remainingSeries - 1);
                int currentLayerSize = Math.Max(1, scatterChart.Elements<C.ScatterChartSeries>().Count());
                int seriesCount = layerIndex == scatterCharts.Count - 1
                    ? remainingSeries
                    : Math.Min(currentLayerSize, remainingSeries - seriesReservedForFutureLayers);
                UpdateScatterChartSeries(scatterChart, data, seriesOffset, seriesCount);
                seriesOffset += seriesCount;
            }

            RemoveUnusedScatterAxes(plotArea);
        }

        private static void UpdateScatterChartSeries(C.ScatterChart scatterChart,
            PowerPointScatterChartData data, int seriesOffset, int seriesCount) {
            List<C.ScatterChartSeries> existingSeries = scatterChart.Elements<C.ScatterChartSeries>().ToList();
            C.ScatterChartSeries? template = existingSeries.LastOrDefault();

            for (int localIndex = 0; localIndex < seriesCount; localIndex++) {
                C.ScatterChartSeries seriesElement;
                if (localIndex < existingSeries.Count) {
                    seriesElement = existingSeries[localIndex];
                } else {
                    seriesElement = template != null
                        ? (C.ScatterChartSeries)template.CloneNode(true)
                        : new C.ScatterChartSeries();
                    seriesElement.RemoveAllChildren<C.Trendline>();
                    InsertSeries(scatterChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                int seriesIndex = seriesOffset + localIndex;
                UpdateSeriesIndexOrder(seriesElement, seriesIndex);
                UpdateScatterSeriesText(seriesElement, seriesIndex, data.Series[seriesIndex].Name);
                UpdateXValues(seriesElement, seriesIndex, data.Series[seriesIndex].XValues);
                UpdateYValues(seriesElement, seriesIndex, data.Series[seriesIndex].YValues);
            }

            for (int localIndex = existingSeries.Count - 1; localIndex >= seriesCount; localIndex--) {
                existingSeries[localIndex].Remove();
            }
        }

        private static void RemoveUnusedScatterAxes(C.PlotArea plotArea) {
            var usedAxisIds = new HashSet<uint>(plotArea.Elements<C.ScatterChart>()
                .SelectMany(chart => chart.Elements<C.AxisId>())
                .Where(axis => axis.Val?.Value != null)
                .Select(axis => axis.Val!.Value));
            foreach (C.ValueAxis axis in plotArea.Elements<C.ValueAxis>().ToList()) {
                uint? axisId = axis.AxisId?.Val?.Value;
                if (axisId.HasValue && !usedAxisIds.Contains(axisId.Value)) axis.Remove();
            }
        }
    }
}
