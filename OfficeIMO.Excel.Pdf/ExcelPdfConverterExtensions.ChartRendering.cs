using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static readonly OfficeChartLayout DefaultExcelPdfChartLayout = new OfficeChartLayout(
            seriesLegendWidthRatio: 0.0001D,
            categoryLegendWidthRatio: 0.0001D);

        private static void AddWorksheetChart(PdfCore.PdfContentBuilder item, WorksheetChartExportData chart, string sheetName, ExcelPdfExportOptions options) {
            ExcelChartInfo snapshot = chart.Snapshot;
            if (string.IsNullOrWhiteSpace(snapshot.Title) && !string.IsNullOrWhiteSpace(snapshot.Name)) {
                item.H2(snapshot.Name, PdfCore.PdfAlign.Left, PdfCore.PdfColor.FromRgb(31, 78, 121));
            }

            bool useStyledRadialLegend = HasStyledRadialPoints(snapshot);
            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(
                CreateOfficeChartSnapshotCore(snapshot, options, preserveWorksheetLegend: useStyledRadialLegend));
            AddChartQualityWarning(options, sheetName, snapshot, rendering.QualityReport);
            item.Drawing(rendering.Drawing, PdfCore.PdfAlign.Left, spacingBefore: 2D, spacingAfter: 6D);
            // The shared chart legend renders point fill, no-fill, hatch, and outline. A table cell
            // can only reproduce a solid fill, so keep the data table uncolored for styled radial charts.
            item.Table(CreateChartLegendRows(snapshot), PdfCore.PdfAlign.Left,
                CreateChartLegendStyle(useStyledRadialLegend ? 0 : GetChartLegendColorCount(snapshot), options.ChartStyle));
        }

        private static void AddChartQualityWarning(ExcelPdfExportOptions options, string sheetName, ExcelChartInfo snapshot, OfficeDrawingQualityReport qualityReport) {
            if (!qualityReport.HasIssues) {
                return;
            }

            AddWarning(
                options,
                sheetName,
                "chart-quality",
                "Exported worksheet chart '" + (string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title) + "' with shared drawing quality warnings: " + FormatQualityIssues(qualityReport));
        }

        private static string FormatQualityIssues(OfficeDrawingQualityReport qualityReport) {
            return string.Join("; ", qualityReport.Issues.Select(issue => issue.ToString()));
        }

        private static OfficeChartSnapshot CreateOfficeChartSnapshot(ExcelChartInfo snapshot, ExcelPdfExportOptions options) {
            return CreateOfficeChartSnapshotCore(snapshot, options, preserveWorksheetLegend: false);
        }

        private static OfficeChartSnapshot CreateOfficeChartSnapshotCore(ExcelChartInfo snapshot, ExcelPdfExportOptions options, bool preserveWorksheetLegend) {
            if (!TryMapChartKind(snapshot.ChartType, out OfficeChartType chartKind)) {
                throw new NotSupportedException("Excel chart type '" + snapshot.ChartType + "' is not supported by the shared OfficeIMO chart renderer.");
            }

            var series = snapshot.Data.Series
                .Select(item => CreateOfficeChartSeries(item, snapshot.ChartType, chartKind, options.ChartStyle != null))
                .ToList();
            var data = new OfficeChartData(snapshot.Data.Categories, series);
            return new OfficeChartSnapshot(
                snapshot.Name,
                snapshot.Title,
                chartKind,
                data,
                PixelsToPoints(snapshot.WidthPixels),
                PixelsToPoints(snapshot.HeightPixels),
                options.ChartStyle ?? snapshot.Style,
                options.ChartLayout ?? (preserveWorksheetLegend || chartKind is OfficeChartType.Pie or OfficeChartType.Doughnut
                    ? snapshot.Layout ?? new OfficeChartLayout()
                    : WithNativeCategoryGeometry(DefaultExcelPdfChartLayout, snapshot.Layout)),
                snapshot.BubbleScalePercent, snapshot.BubbleSizeMode, snapshot.RadialLayout);
        }

        // The PDF chart keeps its worksheet legend table, but bar spacing, point placement and
        // high-low lines still come from the native chart.
        private static OfficeChartLayout WithNativeCategoryGeometry(OfficeChartLayout layout, OfficeChartLayout? native) =>
            native == null ? layout : layout
                .WithBarGeometry(native.BarGapWidthPercent, native.BarOverlapPercent)
                .WithCategoryPointsOnTickMarks(native.CategoryPointsOnTickMarks)
                .WithHighLowLines(native.ShowHighLowLines, native.HighLowLineColor, native.HighLowLineWidth);

        private static OfficeSeries CreateOfficeChartSeries(ExcelChartSeries series, ExcelChartType defaultType, OfficeChartType defaultKind,
            bool callerPalette = false) {
            if (defaultType == ExcelChartType.Bubble) return ExcelRangeImageRenderer.CreateBubbleChartSeries(series, callerPalette);
            ExcelChartType effectiveType = series.ChartType ?? defaultType;
            OfficeChartType? renderKind = null;
            if (effectiveType != defaultType) {
                if (!ExcelRangeImageRenderer.TryMapCompatibleComboSeriesRenderKind(
                        defaultType,
                        effectiveType,
                        out OfficeChartType mappedKind,
                        out _,
                        out string? unsupportedReason)) {
                    throw new NotSupportedException(
                        "Excel combo-chart series '" + series.Name + "' uses chart type '" + effectiveType + "': " + unsupportedReason);
                }

                renderKind = mappedKind == defaultKind ? null : mappedKind;
            }

            return new OfficeSeries(
                series.Name,
                series.Values,
                series.XValues,
                // A caller palette replaces automatic colours, as it did before they were resolved.
                callerPalette && series.HasAutomaticColor ? null : ExcelRangeImageRenderer.ResolveArgb(series.SeriesColorArgb),
                ExcelRangeImageRenderer.ResolvePointColors(series.PointColorArgb),
                series.ShowMarkers,
                connectLine: series.ConnectLine,
                markerSize: series.MarkerSize,
                markerShape: series.MarkerShape,
                markerOutlineColor: ExcelRangeImageRenderer.ResolveArgb(series.MarkerOutlineColorArgb),
                markerOutlineWidth: series.MarkerOutlineWidth,
                strokeWidth: series.SeriesLineWidth,
                strokeDashStyle: series.SeriesLineDashStyle,
                renderKind: renderKind,
                axisGroup: series.AxisGroup)
                .WithSeriesStyle(series.SeriesStyle).WithPointStyles(series.PointStyles)
                .WithPointExplosions(series.PointExplosions);
        }

        private static bool TryMapChartKind(ExcelChartType type, out OfficeChartType kind) {
            switch (type) {
                case ExcelChartType.ColumnClustered:
                case ExcelChartType.Column3DClustered:
                    kind = OfficeChartType.ColumnClustered;
                    return true;
                case ExcelChartType.ColumnStacked:
                case ExcelChartType.Column3DStacked:
                    kind = OfficeChartType.ColumnStacked;
                    return true;
                case ExcelChartType.ColumnStacked100:
                case ExcelChartType.Column3DStacked100:
                    kind = OfficeChartType.ColumnStacked100;
                    return true;
                case ExcelChartType.BarClustered:
                case ExcelChartType.Bar3DClustered:
                    kind = OfficeChartType.BarClustered;
                    return true;
                case ExcelChartType.BarStacked:
                case ExcelChartType.Bar3DStacked:
                    kind = OfficeChartType.BarStacked;
                    return true;
                case ExcelChartType.BarStacked100:
                case ExcelChartType.Bar3DStacked100:
                    kind = OfficeChartType.BarStacked100;
                    return true;
                case ExcelChartType.Line:
                case ExcelChartType.Line3D:
                    kind = OfficeChartType.Line;
                    return true;
                case ExcelChartType.LineStacked:
                    kind = OfficeChartType.LineStacked;
                    return true;
                case ExcelChartType.LineStacked100:
                    kind = OfficeChartType.LineStacked100;
                    return true;
                case ExcelChartType.Area:
                case ExcelChartType.Area3D:
                    kind = OfficeChartType.Area;
                    return true;
                case ExcelChartType.AreaStacked:
                case ExcelChartType.Area3DStacked:
                    kind = OfficeChartType.AreaStacked;
                    return true;
                case ExcelChartType.AreaStacked100:
                case ExcelChartType.Area3DStacked100:
                    kind = OfficeChartType.AreaStacked100;
                    return true;
                case ExcelChartType.Bubble:
                    kind = OfficeChartType.Bubble;
                    return true;
                case ExcelChartType.Scatter:
                    kind = OfficeChartType.Scatter;
                    return true;
                case ExcelChartType.Radar:
                    kind = OfficeChartType.Radar;
                    return true;
                case ExcelChartType.Pie:
                case ExcelChartType.Pie3D:
                case ExcelChartType.PieOfPie:
                case ExcelChartType.BarOfPie:
                    kind = OfficeChartType.Pie;
                    return true;
                case ExcelChartType.Doughnut:
                    kind = OfficeChartType.Doughnut;
                    return true;
                case ExcelChartType.Stock:
                case ExcelChartType.Surface:
                case ExcelChartType.SurfaceWireframe:
                case ExcelChartType.SurfaceContour:
                case ExcelChartType.SurfaceContourWireframe:
                    kind = OfficeChartType.Line;
                    return true;
                default:
                    kind = default;
                    return false;
            }
        }

        private static string[][] CreateChartLegendRows(ExcelChartInfo snapshot) {
            if (IsPieLikeChart(snapshot.ChartType)) {
                return CreatePieChartLegendRows(snapshot);
            }

            var rows = new List<string[]> {
                new[] { "Series", "Values" }
            };

            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                rows.Add(new[] {
                    series.Name,
                    string.Join(", ", series.Values.Select(value => value.ToString("0.##", CultureInfo.InvariantCulture)))
                });
            }

            return rows.ToArray();
        }

        private static string[][] CreatePieChartLegendRows(ExcelChartInfo snapshot) {
            var rows = new List<string[]> {
                new[] { "Category", "Value" }
            };

            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            ExcelChartSeries? values = series.Count > 0 ? series[0] : null;
            for (int i = 0; i < categories.Count; i++) {
                string category = string.IsNullOrWhiteSpace(categories[i])
                    ? "Slice " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    : categories[i];
                rows.Add(new[] {
                    category,
                    values == null ? string.Empty : GetSeriesValue(values, i).ToString("0.##", CultureInfo.InvariantCulture)
                });
            }

            return rows.ToArray();
        }

        private static int GetChartLegendColorCount(ExcelChartInfo snapshot) {
            if (IsPieLikeChart(snapshot.ChartType)) {
                return snapshot.Data.Categories.Count;
            }

            return snapshot.Data.Series.Count;
        }

        private static bool HasStyledRadialPoints(ExcelChartInfo snapshot) {
            if (!IsPieLikeChart(snapshot.ChartType)) return false;
            return snapshot.Data.Series.Any(series =>
                series.PointStyles?.Any(style => style != null) == true ||
                series.PointColorArgb?.Any(color => !string.IsNullOrWhiteSpace(color)) == true);
        }

        private static PdfCore.PdfTableStyle CreateChartLegendStyle(int colorCount, OfficeChartStyle? chartStyle) {
            var style = new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                FontSize = 8.5D,
                HeaderFontSize = 8.5D,
                CellPaddingX = 4D,
                CellPaddingY = 2D,
                BorderColor = PdfCore.PdfColor.FromRgb(203, 213, 225),
                HeaderFill = PdfCore.PdfColor.FromRgb(239, 246, 255),
                ColumnWidthWeights = new List<double> { 0.7D, 1.3D },
                AutoFitColumns = false,
                MaxWidth = 300D,
                SpacingAfter = 6D
            };

            var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            OfficeChartStyle paletteStyle = chartStyle ?? OfficeChartStyle.Default;
            for (int i = 0; i < colorCount; i++) {
                fills[(i + 1, 0)] = PdfCore.PdfColor.FromOfficeColor(paletteStyle.GetSeriesColor(i));
            }
            style.CellFills = fills;
            return style;
        }

        private static bool IsPieLikeChart(ExcelChartType type) {
            return type == ExcelChartType.Pie
                   || type == ExcelChartType.Pie3D
                   || type == ExcelChartType.PieOfPie
                   || type == ExcelChartType.BarOfPie
                   || type == ExcelChartType.Doughnut;
        }

        private static double GetSeriesValue(ExcelChartSeries series, int index) {
            double value = index >= 0 && index < series.Values.Count ? series.Values[index] : 0D;
            return double.IsNaN(value) || double.IsInfinity(value) ? 0D : value;
        }
    }
}
