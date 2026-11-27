using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>Updates the native chart from the shared OfficeIMO chart contract.</summary>
        public PowerPointChart UpdateData(OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(data);
            PowerPointImportedChartReport imported = InspectImportedContent();
            if (AdvancedChartProjections.ContainsKey(imported.Family)) {
                return UpdateImportedData(data);
            }
            if (!TryGetSnapshotForUpdate(out PowerPointChartSnapshot current)) {
                throw new NotSupportedException(
                    "The current chart kind cannot be updated through the shared OfficeIMO chart contract.");
            }
            OfficeChartType chartKind = MapKind(current.ChartKind);
            OfficeOpenXmlChartWriter.ValidateSharedChartData(data, chartKind);

            IOfficeChartPart chartPart = SharedChartPart;
            IOfficeChartEmbeddedPackage? embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
            if (chartKind == OfficeChartType.Bubble && embedded == null) {
                throw new NotSupportedException(
                    "Bubble chart data cannot be updated without an embedded workbook.");
            }
            byte[]? workbookBytes = embedded != null ? OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind) : null;
            Action? preserveBindings = embedded != null ? OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                chartPart, data.Series.Select(series => series.Values.Count).ToArray()) : null;
            PowerPointChartAxisIdGenerator.EnsureSeeded(ChartPartRef.Document);
            OfficeOpenXmlChartWriter.UpdateSharedChartData(chartPart, data, chartKind,
                ColorTheme);
            preserveBindings?.Invoke();

            embedded?.Replace(workbookBytes!);
            MarkChartEdited();
            return this;
        }

        /// <summary>Creates a deterministic plain-text summary suitable for accessibility review or sidecar output.</summary>
        public static string CreateDataSummary(OfficeChartType chartKind, OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(data);
            var builder = new StringBuilder();
            builder.Append("Chart kind: ").Append(chartKind).AppendLine();
            if ((chartKind == OfficeChartType.Scatter || chartKind == OfficeChartType.Bubble) &&
                data.Series.Any(series => series.XValues != null)) {
                AppendNumericPointDataSummary(builder, data, includeBubbleSize: chartKind == OfficeChartType.Bubble);
                return builder.ToString();
            }
            builder.Append("Category");
            foreach (OfficeSeries series in data.Series) {
                builder.Append('\t').Append(CleanSummaryValue(series.Name));
            }
            builder.AppendLine();
            for (int categoryIndex = 0; categoryIndex < data.Categories.Count; categoryIndex++) {
                builder.Append(CleanSummaryValue(data.Categories[categoryIndex]));
                foreach (OfficeSeries series in data.Series) {
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
            foreach (OfficeSeries series in data.Series) {
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
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            OfficeFileCommit.WriteAllBytes(filePath, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(CreateDataSummary()));
            return this;
        }

        /// <summary>Applies native alternative text and optionally includes a plain-text data summary.</summary>
        public PowerPointChart SetAccessibility(string alternativeText, string? dataSummary = null,
            bool includeDataSummary = true) {
            ArgumentException.ThrowIfNullOrWhiteSpace(alternativeText);
            string resolved = dataSummary ?? (includeDataSummary ? CreateDataSummary() : string.Empty);
            AltText = includeDataSummary && !string.IsNullOrWhiteSpace(resolved)
                ? alternativeText.Trim() + Environment.NewLine + Environment.NewLine + "Data summary:" +
                  Environment.NewLine + resolved.Trim()
                : alternativeText.Trim();
            return this;
        }

        /// <summary>Tries to expose the current chart through the shared dependency-free chart contract.</summary>
        public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) =>
            TryGetOfficeSnapshot(null, null, out snapshot, out _, out _);

        /// <summary>Projects the chart at an optional logical extent (point units by default) and reports why an unsupported chart cannot be projected.</summary>
        internal bool TryGetOfficeSnapshot(double? width, double? height, out OfficeChartSnapshot snapshot,
            out string? failureReason, out string? projectionNotice) {
            projectionNotice = null;
            if (!TryGetSnapshot(out PowerPointChartSnapshot powerPointSnapshot, out failureReason)) {
                snapshot = null!;
                failureReason ??= "The chart could not be projected.";
                return false;
            }
            if (powerPointSnapshot.Data.Series.Any(item => item.BubbleSizes != null &&
                (item.XValues == null || item.BubbleSizes.Any(size =>
                    double.IsNaN(size) || double.IsInfinity(size) || size < 0D)))) {
                snapshot = null!;
                failureReason = "The bubble sizes are missing X values or are not finite non-negative numbers.";
                return false;
            }
            snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(powerPointSnapshot,
                width ?? powerPointSnapshot.Width.ToPoints(), height ?? powerPointSnapshot.Height.ToPoints());

            if (HasUnprojectedNativeAxisUnits(snapshot.Layout) ||
                OfficeChartDrawingRenderer.HasUnsupportedAxisUnitBudget(snapshot)) {
                snapshot = null!;
                failureReason = "The chart axis units cannot be projected.";
                return false;
            }
            projectionNotice = powerPointSnapshot.ProjectionNotice;
            return true;
        }

        private bool HasUnprojectedNativeAxisUnits(OfficeChartLayout layout) {
            OfficeXmlNode secondaryAxis = OfficeOpenXmlChartSecondaryAxis.Resolve(PlotAreaNode);
            foreach (OfficeXmlNode axis in ChartSpaceNode.Descendants(CNs, "valAx")) {
                bool secondary = axis == secondaryAxis;
                string? position = axis.ChildVal(CNs, "axPos");
                OfficeXmlNode majorUnit = axis.Child(CNs, "majorUnit"), minorUnit = axis.Child(CNs, "minorUnit");
                if (position == null && (majorUnit.Exists || minorUnit.Exists)) return true;
                bool horizontal = position is "b" or "t";
                double? projectedMajor = secondary ? layout.SecondaryValueAxis?.MajorUnit :
                    horizontal ? layout.HorizontalAxisMajorUnit : layout.VerticalAxisMajorUnit;
                double? projectedMinor = secondary ? layout.SecondaryValueAxis?.MinorUnit :
                    horizontal ? layout.HorizontalAxisMinorUnit : layout.VerticalAxisMinorUnit;
                if (majorUnit.DoubleVal() is double major &&
                    major != projectedMajor ||
                    minorUnit.DoubleVal() is double minor &&
                    minor != projectedMinor) return true;
            }
            return false;
        }

        private OfficeChartStyle? ReadSharedTextStyle(OfficeXmlNode chart) {
            return TryReadSharedTextStyle(chart,
                out OfficeChartStyle? style)
                ? style
                : null;
        }

        private bool TryReadSharedTextStyle(OfficeXmlNode chart, out OfficeChartStyle? style) =>
            new OfficeOpenXmlChartTextReader(GetChartThemeFontScheme()).TryReadSharedTextStyle(chart, out style);

        private static string? ReadChartDefaultTypeface(OfficeXmlNode chart) =>
            OfficeOpenXmlChartTextReader.ReadChartDefaultTypeface(chart);

        private bool TryReadAxisTitleTypeface(OfficeXmlNode chart, string? chartDefaultTypeface, out string? axisTitleFont) =>
            new OfficeOpenXmlChartTextReader(GetChartThemeFontScheme()).TryReadAxisTitleTypeface(chart, chartDefaultTypeface, out axisTitleFont);

        private OfficeXmlNode GetChartThemeFontScheme() => PowerPointChartPart.HostScheme(OwnerPartRef, "fontScheme");

        private static HashSet<uint> GetHiddenLegendSeriesIndexes(OfficeXmlNode chart) {
            var result = new HashSet<uint>();
            foreach (OfficeXmlNode entry in chart.Child(CNs, "legend").Children(CNs, "legendEntry")) {
                if (IsOn(entry.Child(CNs, "delete")) &&
                    entry.Child(CNs, "idx").UIntVal() is uint seriesIndex) {
                    result.Add(seriesIndex);
                }
            }
            return result;
        }

        private static string CleanSummaryValue(string? value) =>
            (value ?? string.Empty).Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');

        private static OfficeChartType MapKind(PowerPointChartSnapshotKind kind) =>
            PowerPointChartSnapshotMapper.MapKind(kind);
    }
}
