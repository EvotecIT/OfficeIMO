using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private static PowerPointChartData? ReadBubbleSeriesData(
            IEnumerable<C.BubbleChartSeries> seriesElements, ColorScheme? colorScheme = null) {
            List<C.BubbleChartSeries> source = seriesElements.ToList();
            if (source.Count == 0) return null;

            var series = new List<PowerPointChartSeries>();
            IReadOnlyList<double>? categoryXValues = null;
            for (int seriesIndex = 0; seriesIndex < source.Count; seriesIndex++) {
                C.BubbleChartSeries element = source[seriesIndex];
                IReadOnlyList<double> xValues =
                    ReadCachedNumbers(element.GetFirstChild<C.XValues>());
                IReadOnlyList<double> yValues =
                    ReadCachedNumbers(element.GetFirstChild<C.YValues>());
                IReadOnlyList<double> bubbleSizes =
                    ReadCachedNumbers(element.GetFirstChild<C.BubbleSize>());
                if (xValues.Count != yValues.Count || xValues.Count != bubbleSizes.Count) {
                    return null;
                }

                int pointCount = xValues.Count;
                if (pointCount == 0) continue;

                IReadOnlyList<double> normalizedY = NormalizeValues(yValues, pointCount);
                IReadOnlyList<double> normalizedSizes = NormalizeValues(bubbleSizes, pointCount);
                if (normalizedY.Count == 0 || normalizedSizes.Count == 0) continue;

                IReadOnlyList<double> normalizedX = xValues.Take(pointCount).ToList();
                categoryXValues ??= normalizedX;
                string name = ReadSeriesName(element);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (seriesIndex + 1).ToString(CultureInfo.InvariantCulture);
                }

                var item = new PowerPointChartSeries(name, normalizedY, normalizedX,
                    PowerPointChartSnapshotKind.Bubble,
                    ReadSeriesColor(element, PowerPointChartSnapshotKind.Bubble, colorScheme),
                    ReadSeriesStrokeWidth(element)) {
                    BubbleSizes = normalizedSizes,
                    PointColors = ReadBubblePointColors(element, pointCount, colorScheme),
                    SourceIndex = element.GetFirstChild<C.Index>()?.Val?.Value
                };
                series.Add(item);
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = categoryXValues.Select(value =>
                value.ToString(CultureInfo.InvariantCulture)).ToList();
            return new PowerPointChartData(categories, series);
        }

        private static IReadOnlyList<OfficeColor?>? ReadBubblePointColors(
            C.BubbleChartSeries series, int pointCount, ColorScheme? colorScheme) {
            var colors = new OfficeColor?[pointCount];
            bool found = false;
            foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
                uint? sourceIndex = point.GetFirstChild<C.Index>()?.Val?.Value;
                if (!sourceIndex.HasValue || sourceIndex.Value >= (uint)pointCount) continue;
                C.ChartShapeProperties? properties =
                    point.GetFirstChild<C.ChartShapeProperties>();
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    properties?.GetFirstChild<SolidFill>(), colorScheme);
                if (!color.HasValue) continue;
                colors[(int)sourceIndex.Value] = color;
                found = true;
            }
            return found ? colors : null;
        }
    }
}
