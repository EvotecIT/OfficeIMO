using System;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderChart(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointChart chart, double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
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

    private static void AddChartQualityWarning(PowerPointToPdfOptions options, int slideNumber, PptCore.PowerPointChartSnapshot snapshot, OfficeDrawingQualityReport qualityReport, double x, double y, double width, double height) {
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

    /// <summary>
    /// Maps the native PowerPoint snapshot into the shared chart contract used by PDF rendering.
    /// </summary>
    internal static OfficeChartSnapshot CreateOfficeChartSnapshot(PptCore.PowerPointChartSnapshot snapshot, double width, double height, PowerPointToPdfOptions options) =>
        PptCore.PowerPointChartSnapshotMapper.ToOfficeSnapshot(snapshot, width, height,
            options.ChartStyle, options.ChartLayout);
}
