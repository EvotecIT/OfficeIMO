using System;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderChart(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointChart chart, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        if (!chart.TryGetSnapshot(out PptCore.PowerPointChartSnapshot snapshot)) {
            AddLayoutWarning(
                options,
                slideNumber,
                "unsupported-chart",
                "Skipped a PowerPoint chart because its cached chart data could not be read into a first-party PDF snapshot.",
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointChart",
                "The PowerPoint chart snapshot could not be read into the shared PDF chart renderer.",
                x,
                y,
                width,
                height);
            return;
        }

        try {
            OfficeChartSnapshot chartSnapshot = CreateOfficeChartSnapshot(snapshot, width, height, options);
            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(chartSnapshot);
            AddChartQualityWarning(options, slideNumber, snapshot, rendering.QualityReport, x, y, width, height);
            canvas.Drawing(
                rendering.Drawing,
                x,
                y,
                width,
                height,
                style: new PdfCore.PdfDrawingStyle {
                    AlternativeText = string.IsNullOrWhiteSpace(snapshot.Title) ? "PowerPoint chart" : snapshot.Title
                },
                rotationAngle: chart.Rotation ?? 0D);
        } catch (Exception ex) {
            AddLayoutWarning(
                options,
                slideNumber,
                "unsupported-chart",
                "Skipped a PowerPoint chart because it could not be rendered as a shared PDF chart: " + ex.Message,
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointChart",
                "The PowerPoint chart could not be rendered by the shared PDF chart renderer.",
                x,
                y,
                width,
                height);
        }
    }

    private static void AddChartQualityWarning(PowerPointPdfSaveOptions options, int slideNumber, PptCore.PowerPointChartSnapshot snapshot, OfficeDrawingQualityReport qualityReport, double x, double y, double width, double height) {
        if (!qualityReport.HasIssues) {
            return;
        }

        AddLayoutWarning(
            options,
            slideNumber,
            "chart-quality",
            "Rendered PowerPoint chart '" + (string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title) + "' with shared drawing quality warnings: " + FormatQualityIssues(qualityReport),
            PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
            "PowerPointChart",
            "The shared PDF chart renderer reported visual quality issues.",
            x,
            y,
            width,
            height);
    }

    private static string FormatQualityIssues(OfficeDrawingQualityReport qualityReport) {
        return string.Join("; ", qualityReport.Issues.Select(issue => issue.ToString()));
    }

    private static OfficeChartSnapshot CreateOfficeChartSnapshot(PptCore.PowerPointChartSnapshot snapshot, double width, double height, PowerPointPdfSaveOptions options) {
        var series = snapshot.Data.Series
            .Select(item => new OfficeChartSeries(item.Name, item.Values, item.XValues))
            .ToList();
        var data = new OfficeChartData(snapshot.Data.Categories, series);
        return new OfficeChartSnapshot(
            snapshot.Name,
            snapshot.Title,
            MapChartKind(snapshot.ChartKind),
            data,
            width,
            height,
            options.ChartStyle,
            options.ChartLayout);
    }

    private static OfficeChartKind MapChartKind(PptCore.PowerPointChartSnapshotKind kind) {
        switch (kind) {
            case PptCore.PowerPointChartSnapshotKind.ClusteredColumn:
                return OfficeChartKind.ColumnClustered;
            case PptCore.PowerPointChartSnapshotKind.StackedColumn:
                return OfficeChartKind.ColumnStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedColumn100:
                return OfficeChartKind.ColumnStacked100;
            case PptCore.PowerPointChartSnapshotKind.ClusteredBar:
                return OfficeChartKind.BarClustered;
            case PptCore.PowerPointChartSnapshotKind.StackedBar:
                return OfficeChartKind.BarStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedBar100:
                return OfficeChartKind.BarStacked100;
            case PptCore.PowerPointChartSnapshotKind.Line:
                return OfficeChartKind.Line;
            case PptCore.PowerPointChartSnapshotKind.StackedLine:
                return OfficeChartKind.LineStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedLine100:
                return OfficeChartKind.LineStacked100;
            case PptCore.PowerPointChartSnapshotKind.Area:
                return OfficeChartKind.Area;
            case PptCore.PowerPointChartSnapshotKind.StackedArea:
                return OfficeChartKind.AreaStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedArea100:
                return OfficeChartKind.AreaStacked100;
            case PptCore.PowerPointChartSnapshotKind.Radar:
                return OfficeChartKind.Radar;
            case PptCore.PowerPointChartSnapshotKind.Scatter:
                return OfficeChartKind.Scatter;
            case PptCore.PowerPointChartSnapshotKind.Pie:
                return OfficeChartKind.Pie;
            case PptCore.PowerPointChartSnapshotKind.Doughnut:
                return OfficeChartKind.Doughnut;
            default:
                throw new NotSupportedException("PowerPoint chart kind '" + kind + "' is not supported by the shared PDF chart renderer.");
        }
    }
}
