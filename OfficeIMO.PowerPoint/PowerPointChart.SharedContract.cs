using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
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
            PowerPointImportedChartReport imported = InspectImportedContent();
            if (AdvancedChartProjections.ContainsKey(imported.Family)) {
                return UpdateImportedData(data);
            }
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

        private OfficeChartStyle? ReadSharedTextStyle(C.Chart chart) {
            return TryReadSharedTextStyle(chart,
                out OfficeChartStyle? style)
                ? style
                : null;
        }

        private bool TryReadSharedTextStyle(C.Chart chart,
            out OfficeChartStyle? style) {
            string? chartDefaultTypeface = ReadChartDefaultTypeface(chart);
            OpenXmlElement[] bodyTextAreas = chart.Descendants()
                .Where(IsRelevantBodyTextArea)
                .ToArray();
            string?[] bodyFonts = bodyTextAreas
                .Select(textArea => ReadBodyTypeface(textArea,
                    chartDefaultTypeface))
                .ToArray();
            if (bodyFonts.Length > 0
                && (bodyFonts.Any(string.IsNullOrWhiteSpace)
                    || bodyFonts.Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count() != 1)) {
                style = null;
                return false;
            }
            string? bodyFont = bodyFonts.Length == 0
                ? null
                : bodyFonts[0];
            C.Title? title = chart.GetFirstChild<C.Title>();
            string? titleFont = ReadTitleTypeface(title,
                chartDefaultTypeface);
            if (title != null
                && title.Descendants<A.Text>()
                    .Any(text => !string.IsNullOrEmpty(text.Text))
                && string.IsNullOrWhiteSpace(titleFont)) {
                style = null;
                return false;
            }
            if (!TryReadAxisTitleTypeface(chart, chartDefaultTypeface,
                    out _)) {
                style = null;
                return false;
            }
            style = bodyFont == null && titleFont == null
                ? null
                : new OfficeChartStyle(fontFamily: bodyFont,
                    titleFontFamily: titleFont);
            return true;
        }

        private static string? ReadChartDefaultTypeface(C.Chart chart) =>
            chart.Parent?.GetFirstChild<C.TextProperties>()?
                .Descendants<A.LatinFont>()
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

        private bool TryReadAxisTitleTypeface(C.Chart chart,
            string? chartDefaultTypeface, out string? axisTitleFont) {
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            C.Title[] titles = plotArea == null
                ? Array.Empty<C.Title>()
                : plotArea.Elements()
                    .Where(axis => axis is C.CategoryAxis
                        or C.DateAxis or C.ValueAxis or C.SeriesAxis)
                    .Select(axis => axis.GetFirstChild<C.Title>())
                    .Where(title => title != null
                        && title.Descendants<A.Text>()
                            .Any(text => !string.IsNullOrEmpty(text.Text)))
                    .Cast<C.Title>()
                    .ToArray();
            string?[] fonts = titles.Select(title =>
                    ReadTitleTypeface(title, chartDefaultTypeface))
                .ToArray();
            if (fonts.Any(string.IsNullOrWhiteSpace)) {
                axisTitleFont = null;
                return false;
            }
            string[] distinct = fonts.Select(font => font!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            axisTitleFont = distinct.Length == 1 ? distinct[0] : null;
            return distinct.Length <= 1;
        }

        private string? ReadBodyTypeface(
            OpenXmlElement textArea, string? chartDefaultTypeface) {
            string?[] typefaces = textArea.Descendants<C.TextProperties>()
                .Where(properties => !properties.Ancestors<C.Title>().Any())
                .SelectMany(properties => properties.Descendants<A.LatinFont>())
                .Select(font => font.Typeface?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToArray();
            if (typefaces.Length == 0) {
                return ResolveChartTypeface(chartDefaultTypeface,
                    useMajorWhenMissing: false);
            }
            string?[] resolved = typefaces.Select(value =>
                    ResolveChartTypeface(value, useMajorWhenMissing: false))
                .ToArray();
            if (resolved.Any(string.IsNullOrWhiteSpace)) return null;
            string[] explicitFonts = resolved.Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            return explicitFonts.Length == 1 ? explicitFonts[0] : null;
        }

        private string? ReadTitleTypeface(C.Title? title,
            string? chartDefaultTypeface) {
            if (title == null) return null;
            string? defaultTypeface = title.Elements<C.TextProperties>()
                .SelectMany(properties => properties.Descendants<A.LatinFont>())
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                ?? chartDefaultTypeface;
            OpenXmlElement[] textRuns = title.Descendants<C.RichText>()
                .SelectMany(richText => richText.Descendants()
                    .Where(element => element is A.Run or A.Field))
                .Where(run => run.GetFirstChild<A.Text>() != null)
                .ToArray();
            if (textRuns.Length == 0) {
                return ResolveChartTypeface(defaultTypeface,
                    useMajorWhenMissing: true);
            }

            string?[] resolvedFonts = textRuns.Select(run =>
                    ResolveChartTypeface(ReadTitleRunTypeface(run)
                            ?? defaultTypeface,
                        useMajorWhenMissing: true))
                .ToArray();
            if (resolvedFonts.Any(string.IsNullOrWhiteSpace)) {
                return null;
            }
            string[] explicitFonts = resolvedFonts.Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            return explicitFonts.Length == 1 ? explicitFonts[0] : null;
        }

        private string? ResolveChartTypeface(string? typeface,
            bool useMajorWhenMissing) {
            if (!string.IsNullOrWhiteSpace(typeface)
                && !IsThemeFontToken(typeface)) {
                return typeface;
            }
            A.FontScheme? scheme = GetChartThemeFontScheme();
            if (scheme == null) return null;
            if (string.IsNullOrWhiteSpace(typeface)) {
                return useMajorWhenMissing
                    ? scheme.MajorFont?.LatinFont?.Typeface?.Value
                    : scheme.MinorFont?.LatinFont?.Typeface?.Value;
            }
            if (typeface!.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase)) {
                return scheme.MajorFont?.LatinFont?.Typeface?.Value;
            }
            if (typeface.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase)) {
                return scheme.MinorFont?.LatinFont?.Typeface?.Value;
            }
            return null;
        }

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

        private static string? ReadTitleRunTypeface(OpenXmlElement run) {
            string? runTypeface = run.GetFirstChild<A.RunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(runTypeface)) return runTypeface;
            A.Paragraph? paragraph = run.Ancestors<A.Paragraph>()
                .FirstOrDefault();
            string? paragraphTypeface = paragraph?
                .GetFirstChild<A.ParagraphProperties>()?
                .GetFirstChild<A.DefaultRunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(paragraphTypeface)) {
                return paragraphTypeface;
            }

            int level = paragraph?.GetFirstChild<A.ParagraphProperties>()?
                .Level?.Value ?? 0;
            A.ListStyle? listStyle = run.Ancestors<C.RichText>()
                .FirstOrDefault()?.GetFirstChild<A.ListStyle>();
            OpenXmlCompositeElement? levelProperties = listStyle?.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .FirstOrDefault(element => string.Equals(element.LocalName,
                    $"lvl{level + 1}pPr", StringComparison.Ordinal));
            return levelProperties?.GetFirstChild<A.DefaultRunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
        }

        private static bool IsRelevantBodyTextArea(OpenXmlElement element) {
            if (element is C.Legend or C.CategoryAxis or C.ValueAxis
                or C.DateAxis or C.SeriesAxis or C.DisplayUnitsLabel
                or C.DataTable or C.TrendlineLabel) {
                return true;
            }
            if (element is not C.DataLabels labels) return false;
            return labels.GetFirstChild<C.ShowValue>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowBubbleSize>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowLegendKey>()?.Val?.Value == true;
        }

        private static bool IsThemeFontToken(string? typeface) =>
            !string.IsNullOrEmpty(typeface)
            && typeface![0] == '+';

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
