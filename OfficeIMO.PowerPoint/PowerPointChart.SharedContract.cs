using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Core.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>Updates the native chart from the shared OfficeIMO chart contract.</summary>
        public PowerPointChart UpdateData(OfficeChartData data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            PowerPointImportedChartReport imported = InspectImportedContent();
            if (AdvancedChartProjections.ContainsKey(imported.Family)) {
                return UpdateImportedData(data);
            }
            if (!TryGetSnapshotForUpdate(out PowerPointChartSnapshot current)) {
                throw new NotSupportedException(
                    "The current chart kind cannot be updated through the shared OfficeIMO chart contract.");
            }
            OfficeChartKind chartKind = MapKind(current.ChartKind);
            OfficeOpenXmlChartWriter.ValidateSharedChartData(data, chartKind);

            ChartPart chartPart = GetChartPart();
            EmbeddedPackagePart? embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
            if (chartKind == OfficeChartKind.Bubble && embedded == null) {
                throw new NotSupportedException(
                    "Bubble chart data cannot be updated without an embedded workbook.");
            }
            byte[]? workbookBytes = embedded != null ? OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind) : null;
            Action? preserveBindings = embedded != null ? OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(chartPart) : null;
            OfficeOpenXmlChartWriter.UpdateSharedChartData(chartPart, data, chartKind);
            preserveBindings?.Invoke();

            if (embedded != null) {
                using var stream = new MemoryStream(workbookBytes!);
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
            if (powerPointSnapshot.Data.Series.Any(item => item.BubbleSizes != null &&
                (item.XValues == null || item.BubbleSizes.Any(size =>
                    double.IsNaN(size) || double.IsInfinity(size) || size < 0D)))) {
                snapshot = null!;
                return false;
            }
            snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(powerPointSnapshot,
                powerPointSnapshot.WidthPoints, powerPointSnapshot.HeightPoints);
            return true;
        }

        private OfficeChartStyle? ReadSharedTextStyle(C.Chart chart) {
            return TryReadSharedTextStyle(chart,
                out OfficeChartStyle? style)
                ? style
                : null;
        }

        private bool TryReadSharedTextStyle(C.Chart chart, out OfficeChartStyle? style) =>
            new OfficeOpenXmlChartTextReader(GetChartThemeFontScheme()).TryReadSharedTextStyle(chart, out style);

        private static string? ReadChartDefaultTypeface(C.Chart chart) =>
            OfficeOpenXmlChartTextReader.ReadChartDefaultTypeface(chart);

        private bool TryReadAxisTitleTypeface(C.Chart chart, string? chartDefaultTypeface, out string? axisTitleFont) =>
            new OfficeOpenXmlChartTextReader(GetChartThemeFontScheme()).TryReadAxisTitleTypeface(chart, chartDefaultTypeface, out axisTitleFont);
        private A.FontScheme? GetChartThemeFontScheme() {
            if (_ownerPart is SlidePart slidePart) {
                return slidePart.ThemeOverridePart?.ThemeOverride?.FontScheme
                    ?? slidePart.SlideLayoutPart?.ThemeOverridePart?
                        .ThemeOverride?.FontScheme
                    ?? slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?
                        .Theme?.ThemeElements?.FontScheme;
            }
            if (_ownerPart is SlideLayoutPart layoutPart) {
                return layoutPart.ThemeOverridePart?.ThemeOverride?.FontScheme
                    ?? layoutPart.SlideMasterPart?.ThemePart?.Theme?
                        .ThemeElements?.FontScheme;
            }
            if (_ownerPart is SlideMasterPart masterPart) {
                return masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            }
            if (_ownerPart is NotesSlidePart notesPart) {
                return notesPart.ThemeOverridePart?.ThemeOverride?.FontScheme
                    ?? notesPart.NotesMasterPart?.ThemePart?.Theme?
                        .ThemeElements?.FontScheme;
            }
            if (_ownerPart is NotesMasterPart notesMasterPart) {
                return notesMasterPart.ThemePart?.Theme?.ThemeElements?
                    .FontScheme;
            }
            return (_ownerPart as HandoutMasterPart)?.ThemePart?.Theme?
                .ThemeElements?.FontScheme;
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

        private static OfficeChartKind MapKind(PowerPointChartSnapshotKind kind) =>
            PowerPointChartSnapshotMapper.MapKind(kind);
    }
}
