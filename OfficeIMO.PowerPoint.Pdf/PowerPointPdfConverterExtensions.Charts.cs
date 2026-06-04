using System;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderChart(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointChart chart, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        if (!chart.TryGetSnapshot(out PptCore.PowerPointChartSnapshot snapshot)) {
            AddWarning(options, slideNumber, "unsupported-chart", "Skipped a PowerPoint chart because its cached chart data could not be read into a first-party PDF snapshot.");
            return;
        }

        try {
            OfficeChartSnapshot chartSnapshot = CreateOfficeChartSnapshot(snapshot, width, height, options);
            OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(chartSnapshot);
            AddChartQualityWarning(options, slideNumber, snapshot, rendering.QualityReport);
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
            AddWarning(options, slideNumber, "unsupported-chart", "Skipped a PowerPoint chart because it could not be rendered as a shared PDF chart: " + ex.Message);
        }
    }

    private static void AddChartQualityWarning(PowerPointPdfSaveOptions options, int slideNumber, PptCore.PowerPointChartSnapshot snapshot, OfficeDrawingQualityReport qualityReport) {
        if (!qualityReport.HasIssues) {
            return;
        }

        AddWarning(
            options,
            slideNumber,
            "chart-quality",
            "Rendered PowerPoint chart '" + (string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title) + "' with shared drawing quality warnings: " + FormatQualityIssues(qualityReport));
    }

    private static string FormatQualityIssues(OfficeDrawingQualityReport qualityReport) {
        return string.Join("; ", qualityReport.Issues.Select(issue => issue.ToString()));
    }

    private static OfficeChartSnapshot CreateOfficeChartSnapshot(PptCore.PowerPointChartSnapshot snapshot, double width, double height, PowerPointPdfSaveOptions options) {
        var series = snapshot.Data.Series
            .Select(item => new OfficeChartSeries(item.Name, item.Values))
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
            case PptCore.PowerPointChartSnapshotKind.Line:
                return OfficeChartKind.Line;
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
