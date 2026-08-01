using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>Updates the native chart from the shared OfficeIMO chart contract.</summary>
        public PowerPointChart UpdateData(OfficeChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (!TryGetSnapshotForUpdate(out PowerPointChartSnapshot current)) {
                throw new NotSupportedException(
                    "The current chart kind cannot be updated through the shared OfficeIMO chart contract.");
            }
            OfficeChartKind chartKind = MapKind(current.ChartKind);
            PowerPointUtils.ValidateSharedChartData(data, chartKind);

            ChartPart chartPart = GetChartPart();
            EmbeddedPackagePart? embedded = chartPart
                .GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
            if (chartKind == OfficeChartKind.Bubble && embedded == null) {
                throw new NotSupportedException(
                    "Bubble chart data cannot be updated without an embedded workbook.");
            }
            PowerPointUtils.UpdateSharedChartData(chartPart, data, chartKind);

            if (embedded != null) {
                byte[] workbookBytes = chartKind switch {
                    OfficeChartKind.Scatter =>
                        PowerPointUtils.BuildChartWorkbook(PowerPointUtils.ToPowerPointScatterChartData(data)),
                    OfficeChartKind.Bubble => PowerPointUtils.BuildBubbleChartWorkbook(data),
                    _ => PowerPointUtils.BuildChartWorkbook(PowerPointUtils.ToPowerPointChartData(data))
                };
                using var stream = new MemoryStream(workbookBytes);
                embedded.FeedData(stream);
            }
            Save();
            return this;
        }

        /// <summary>Creates a deterministic plain-text summary suitable for accessibility review or sidecar output.</summary>
        public static string CreateDataSummary(OfficeChartKind chartKind, OfficeChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            var builder = new StringBuilder();
            builder.Append("Chart kind: ").Append(chartKind).AppendLine();
            if ((chartKind == OfficeChartKind.Scatter || chartKind == OfficeChartKind.Bubble) &&
                data.Series.Any(series => series.XValues != null)) {
                AppendNumericPointDataSummary(builder, data, includeBubbleSize: chartKind == OfficeChartKind.Bubble);
                return builder.ToString();
            }
            builder.Append("Category");
            foreach (OfficeChartSeries series in data.Series) {
                builder.Append('\t').Append(CleanSummaryValue(series.Name));
            }
            builder.AppendLine();
            for (int categoryIndex = 0; categoryIndex < data.Categories.Count; categoryIndex++) {
                builder.Append(CleanSummaryValue(data.Categories[categoryIndex]));
                foreach (OfficeChartSeries series in data.Series) {
                    builder.Append('\t');
                    if (categoryIndex < series.Values.Count) {
                        builder.Append(series.Values[categoryIndex].ToString("G", CultureInfo.InvariantCulture));
                    }
                }
                if (categoryIndex + 1 < data.Categories.Count) builder.AppendLine();
            }
            return builder.ToString();
        }

        private static void AppendNumericPointDataSummary(StringBuilder builder, OfficeChartData data,
            bool includeBubbleSize) {
            builder.AppendLine(includeBubbleSize ? "Series\tX\tY\tSize" : "Series\tX\tY");
            bool firstPoint = true;
            foreach (OfficeChartSeries series in data.Series) {
                for (int pointIndex = 0; pointIndex < series.Values.Count; pointIndex++) {
                    if (!firstPoint) builder.AppendLine();
                    firstPoint = false;
                    builder.Append(CleanSummaryValue(series.Name)).Append('\t');
                    if (series.XValues != null && pointIndex < series.XValues.Count) {
                        builder.Append(series.XValues[pointIndex].ToString("G", CultureInfo.InvariantCulture));
                    } else if (pointIndex < data.Categories.Count) {
                        builder.Append(CleanSummaryValue(data.Categories[pointIndex]));
                    }
                    builder.Append('\t')
                        .Append(series.Values[pointIndex].ToString("G", CultureInfo.InvariantCulture));
                    if (includeBubbleSize) {
                        builder.Append('\t');
                        if (series.BubbleSizes != null && pointIndex < series.BubbleSizes.Count) {
                            builder.Append(series.BubbleSizes[pointIndex]
                                .ToString("G", CultureInfo.InvariantCulture));
                        }
                    }
                }
            }
        }

        /// <summary>Creates a deterministic plain-text data summary from the current native chart.</summary>
        public string CreateDataSummary() {
            if (!TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot)) {
                throw new NotSupportedException("The current chart cannot be represented by the shared chart snapshot contract.");
            }
            return CreateDataSummary(snapshot.ChartKind, snapshot.Data);
        }

        /// <summary>Saves the current chart's plain-text data summary as a UTF-8 sidecar.</summary>
        public PowerPointChart SaveDataSummary(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            OfficeFileCommit.WriteAllBytes(filePath, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(CreateDataSummary()));
            return this;
        }

        /// <summary>Applies native alternative text and optionally includes a plain-text data summary.</summary>
        public PowerPointChart SetAccessibility(string alternativeText, string? dataSummary = null,
            bool includeDataSummary = true) {
            if (string.IsNullOrWhiteSpace(alternativeText)) {
                throw new ArgumentException("Alternative text cannot be empty.", nameof(alternativeText));
            }
            string resolved = dataSummary ?? (includeDataSummary ? CreateDataSummary() : string.Empty);
            AltText = includeDataSummary && !string.IsNullOrWhiteSpace(resolved)
                ? alternativeText.Trim() + Environment.NewLine + Environment.NewLine + "Data summary:" +
                  Environment.NewLine + resolved.Trim()
                : alternativeText.Trim();
            return this;
        }

        /// <summary>Tries to expose the current chart through the shared dependency-free chart contract.</summary>
        public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) {
            if (!TryGetSnapshot(out PowerPointChartSnapshot powerPointSnapshot)) {
                snapshot = null!;
                return false;
            }
            OfficeChartKind kind = MapKind(powerPointSnapshot.ChartKind);
            var series = new List<OfficeChartSeries>(powerPointSnapshot.Data.Series.Count);
            for (int seriesIndex = 0; seriesIndex < powerPointSnapshot.Data.Series.Count; seriesIndex++) {
                PowerPointChartSeries item = powerPointSnapshot.Data.Series[seriesIndex];
                if (item.BubbleSizes != null) {
                    if (item.XValues == null || item.BubbleSizes.Any(size =>
                            double.IsNaN(size) || double.IsInfinity(size) || size < 0D)) {
                        snapshot = null!;
                        return false;
                    }
                    series.Add(OfficeChartSeries.CreateBubble(item.Name, item.XValues!,
                        item.Values, item.BubbleSizes, item.Color, item.PointColors,
                        showInLegend: item.ShowInLegend,
                        markerOutlineColor: item.StrokeColor ?? item.Color,
                        markerOutlineWidth: item.StrokeWidth,
                        showMarkerOutline: item.ShowStroke));
                } else {
                    series.Add(new OfficeChartSeries(item.Name, item.Values, item.XValues, item.Color,
                        pointColors: null, showMarkers: true,
                        showInLegend: item.ShowInLegend, connectLine: true,
                        strokeWidth: item.StrokeWidth,
                        renderKind: item.ChartKind.HasValue ? MapKind(item.ChartKind.Value) : null,
                        axisGroup: item.AxisGroup));
                }
            }
            var data = new OfficeChartData(powerPointSnapshot.Data.Categories, series);
            snapshot = new OfficeChartSnapshot(powerPointSnapshot.Name, powerPointSnapshot.Title, kind, data,
                powerPointSnapshot.WidthPoints, powerPointSnapshot.HeightPoints,
                style: powerPointSnapshot.Style,
                layout: powerPointSnapshot.Layout,
                bubbleScalePercent: powerPointSnapshot.BubbleScalePercent,
                bubbleSizeMode: powerPointSnapshot.BubbleSizeMode);
            return true;
        }

        private static OfficeChartStyle? ReadSharedTextStyle(C.Chart chart) {
            string? bodyFont = chart.Descendants<C.TextProperties>()
                .SelectMany(properties => properties.Descendants<A.LatinFont>())
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            string? titleFont = chart.GetFirstChild<C.Title>()?
                .Descendants<A.LatinFont>()
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            return bodyFont == null && titleFont == null
                ? null
                : new OfficeChartStyle(fontFamily: bodyFont,
                    titleFontFamily: titleFont);
        }

        private static HashSet<uint> GetHiddenLegendSeriesIndexes(C.Chart chart) {
            var result = new HashSet<uint>();
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            if (legend == null) return result;
            foreach (C.LegendEntry entry in legend.Elements<C.LegendEntry>()) {
                C.Delete? delete = entry.GetFirstChild<C.Delete>();
                if (delete != null && delete.Val?.Value != false &&
                    entry.Index?.Val?.Value is uint seriesIndex) {
                    result.Add(seriesIndex);
                }
            }
            return result;
        }

        private static string CleanSummaryValue(string? value) =>
            (value ?? string.Empty).Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');

        private static OfficeChartKind MapKind(PowerPointChartSnapshotKind kind) {
            switch (kind) {
                case PowerPointChartSnapshotKind.ClusteredColumn: return OfficeChartKind.ColumnClustered;
                case PowerPointChartSnapshotKind.StackedColumn: return OfficeChartKind.ColumnStacked;
                case PowerPointChartSnapshotKind.StackedColumn100: return OfficeChartKind.ColumnStacked100;
                case PowerPointChartSnapshotKind.ClusteredBar: return OfficeChartKind.BarClustered;
                case PowerPointChartSnapshotKind.StackedBar: return OfficeChartKind.BarStacked;
                case PowerPointChartSnapshotKind.StackedBar100: return OfficeChartKind.BarStacked100;
                case PowerPointChartSnapshotKind.Line: return OfficeChartKind.Line;
                case PowerPointChartSnapshotKind.StackedLine: return OfficeChartKind.LineStacked;
                case PowerPointChartSnapshotKind.StackedLine100: return OfficeChartKind.LineStacked100;
                case PowerPointChartSnapshotKind.Area: return OfficeChartKind.Area;
                case PowerPointChartSnapshotKind.StackedArea: return OfficeChartKind.AreaStacked;
                case PowerPointChartSnapshotKind.StackedArea100: return OfficeChartKind.AreaStacked100;
                case PowerPointChartSnapshotKind.Scatter: return OfficeChartKind.Scatter;
                case PowerPointChartSnapshotKind.Bubble: return OfficeChartKind.Bubble;
                case PowerPointChartSnapshotKind.Radar: return OfficeChartKind.Radar;
                case PowerPointChartSnapshotKind.Pie: return OfficeChartKind.Pie;
                case PowerPointChartSnapshotKind.Doughnut: return OfficeChartKind.Doughnut;
                default: return OfficeChartKind.ColumnClustered;
            }
        }
    }
}
