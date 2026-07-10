using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>Creates a deterministic plain-text summary suitable for accessibility review or sidecar output.</summary>
        public static string CreateDataSummary(OfficeChartKind chartKind, OfficeChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            var builder = new StringBuilder();
            builder.Append("Chart kind: ").Append(chartKind).AppendLine();
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
            File.WriteAllText(filePath, CreateDataSummary(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
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
            foreach (PowerPointChartSeries item in powerPointSnapshot.Data.Series) {
                series.Add(new OfficeChartSeries(item.Name, item.Values, item.XValues, item.Color,
                    pointColors: null, showMarkers: true, showInLegend: true, connectLine: true,
                    strokeWidth: item.StrokeWidth,
                    renderKind: item.ChartKind.HasValue ? MapKind(item.ChartKind.Value) : null,
                    axisGroup: item.AxisGroup));
            }
            var data = new OfficeChartData(powerPointSnapshot.Data.Categories, series);
            snapshot = new OfficeChartSnapshot(powerPointSnapshot.Name, powerPointSnapshot.Title, kind, data,
                powerPointSnapshot.WidthPoints, powerPointSnapshot.HeightPoints);
            return true;
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
                case PowerPointChartSnapshotKind.Radar: return OfficeChartKind.Radar;
                case PowerPointChartSnapshotKind.Pie: return OfficeChartKind.Pie;
                case PowerPointChartSnapshotKind.Doughnut: return OfficeChartKind.Doughnut;
                default: return OfficeChartKind.ColumnClustered;
            }
        }
    }
}
