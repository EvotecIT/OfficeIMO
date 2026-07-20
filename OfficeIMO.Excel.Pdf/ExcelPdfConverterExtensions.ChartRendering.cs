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

        private static void AddWorksheetChart(PdfCore.PdfItemCompose item, WorksheetChartExportData chart, string sheetName, ExcelPdfSaveOptions options) {
            ExcelChartSnapshot snapshot = chart.Snapshot;
            if (string.IsNullOrWhiteSpace(snapshot.Title) && !string.IsNullOrWhiteSpace(snapshot.Name)) {
                item.H2(snapshot.Name, PdfCore.PdfAlign.Left, PdfCore.PdfColor.FromRgb(31, 78, 121));
            }

            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(CreateOfficeChartSnapshot(snapshot, options));
            AddChartQualityWarning(options, sheetName, snapshot, rendering.QualityReport);
            item.Drawing(rendering.Drawing, PdfCore.PdfAlign.Left, spacingBefore: 2D, spacingAfter: 6D);
            item.Table(CreateChartLegendRows(snapshot), PdfCore.PdfAlign.Left, CreateChartLegendStyle(GetChartLegendColorCount(snapshot), options.ChartStyle));
        }

        private static void AddChartQualityWarning(ExcelPdfSaveOptions options, string sheetName, ExcelChartSnapshot snapshot, OfficeDrawingQualityReport qualityReport) {
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

        private static bool HasMixedSeriesChartTypes(ExcelChartSnapshot snapshot) {
            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                if (series.ChartType.HasValue && series.ChartType.Value != snapshot.ChartType) {
                    return true;
                }
            }

            return false;
        }

        private static OfficeChartSnapshot CreateOfficeChartSnapshot(ExcelChartSnapshot snapshot, ExcelPdfSaveOptions options) {
            return CreateOfficeChartSnapshotCore(snapshot, options, preserveWorksheetLegend: false);
        }

        private static OfficeChartSnapshot CreateOfficeChartSnapshotCore(ExcelChartSnapshot snapshot, ExcelPdfSaveOptions options, bool preserveWorksheetLegend) {
            if (HasMixedSeriesChartTypes(snapshot)) {
                throw new NotSupportedException("Excel chart '" + GetChartDisplayName(snapshot) + "' uses mixed per-series chart types, which are not supported by the shared OfficeIMO chart renderer yet.");
            }

            if (!TryMapChartKind(snapshot.ChartType, out OfficeChartKind chartKind)) {
                throw new NotSupportedException("Excel chart type '" + snapshot.ChartType + "' is not supported by the shared OfficeIMO chart renderer.");
            }

            var series = snapshot.Data.Series
                .Select(item => new OfficeChartSeries(item.Name, item.Values))
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
                options.ChartLayout ?? (preserveWorksheetLegend ? snapshot.Layout ?? new OfficeChartLayout() : DefaultExcelPdfChartLayout));
        }

        private static bool TryMapChartKind(ExcelChartType type, out OfficeChartKind kind) {
            switch (type) {
                case ExcelChartType.ColumnClustered:
                case ExcelChartType.Column3DClustered:
                    kind = OfficeChartKind.ColumnClustered;
                    return true;
                case ExcelChartType.ColumnStacked:
                case ExcelChartType.Column3DStacked:
                    kind = OfficeChartKind.ColumnStacked;
                    return true;
                case ExcelChartType.ColumnStacked100:
                case ExcelChartType.Column3DStacked100:
                    kind = OfficeChartKind.ColumnStacked100;
                    return true;
                case ExcelChartType.BarClustered:
                case ExcelChartType.Bar3DClustered:
                    kind = OfficeChartKind.BarClustered;
                    return true;
                case ExcelChartType.BarStacked:
                case ExcelChartType.Bar3DStacked:
                    kind = OfficeChartKind.BarStacked;
                    return true;
                case ExcelChartType.BarStacked100:
                case ExcelChartType.Bar3DStacked100:
                    kind = OfficeChartKind.BarStacked100;
                    return true;
                case ExcelChartType.Line:
                case ExcelChartType.Line3D:
                    kind = OfficeChartKind.Line;
                    return true;
                case ExcelChartType.LineStacked:
                    kind = OfficeChartKind.LineStacked;
                    return true;
                case ExcelChartType.LineStacked100:
                    kind = OfficeChartKind.LineStacked100;
                    return true;
                case ExcelChartType.Area:
                case ExcelChartType.Area3D:
                    kind = OfficeChartKind.Area;
                    return true;
                case ExcelChartType.AreaStacked:
                case ExcelChartType.Area3DStacked:
                    kind = OfficeChartKind.AreaStacked;
                    return true;
                case ExcelChartType.AreaStacked100:
                case ExcelChartType.Area3DStacked100:
                    kind = OfficeChartKind.AreaStacked100;
                    return true;
                case ExcelChartType.Scatter:
                    kind = OfficeChartKind.Scatter;
                    return true;
                case ExcelChartType.Radar:
                    kind = OfficeChartKind.Radar;
                    return true;
                case ExcelChartType.Pie:
                case ExcelChartType.Pie3D:
                case ExcelChartType.PieOfPie:
                case ExcelChartType.BarOfPie:
                    kind = OfficeChartKind.Pie;
                    return true;
                case ExcelChartType.Doughnut:
                    kind = OfficeChartKind.Doughnut;
                    return true;
                default:
                    kind = default;
                    return false;
            }
        }

        private static string[][] CreateChartLegendRows(ExcelChartSnapshot snapshot) {
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

        private static string[][] CreatePieChartLegendRows(ExcelChartSnapshot snapshot) {
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

        private static int GetChartLegendColorCount(ExcelChartSnapshot snapshot) {
            if (IsPieLikeChart(snapshot.ChartType)) {
                return snapshot.Data.Categories.Count;
            }

            return snapshot.Data.Series.Count;
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
