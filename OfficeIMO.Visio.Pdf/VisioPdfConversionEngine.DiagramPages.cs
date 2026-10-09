using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static partial class VisioPdfConversionEngine {
    private static PdfDocumentConversionResult ConvertDiagramPages(VisioDocument source, VisioToPdfOptions options, CancellationToken token) {
        VisioDrawingOptions drawingOptions = options.DrawingOptions?.Clone() ?? new VisioDrawingOptions();
        if (options.SourceName != null) drawingOptions.SourceName = options.SourceName;
        var report = new PdfConversionReport();
        PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfOptions();
        pdfOptions.ReportDiagnosticsTo(report, "OfficeIMO.Visio.Pdf");
        var profile = new OfficeRenderingProfile(
            "visio-diagram-pages",
            fonts: drawingOptions.Fonts,
            textShapingProvider: drawingOptions.TextShapingProvider,
            textShapingLanguage: drawingOptions.TextShapingLanguage);
        pdfOptions.UseRenderingProfile(profile, OfficeRenderingProfileApplyMode.Overlay);
        drawingOptions.LayoutMetrics = PdfWriter.CreateDrawingTextMetrics(pdfOptions, token);
        OfficeConversionResult<IReadOnlyList<OfficeDrawing>, VisioDrawingConversionReport> projected = source.ToDrawings(drawingOptions, token);
        PdfDocument pdf = PdfDocument.Create(pdfOptions);
        foreach (OfficeDrawing drawing in projected.Value) {
            token.ThrowIfCancellationRequested();
            pdf.Compose(document => document.Page(page => {
                page.Size(drawing.Width, drawing.Height).Margin(0);
                page.Canvas(canvas => {
                    if (drawing.Elements.Count > 0) canvas.Drawing(drawing, 0, 0, drawing.Width, drawing.Height);
                });
            }));
        }
        token.ThrowIfCancellationRequested();
        return new PdfDocumentConversionResult(pdf, report).WithSourceConversionReport(projected.Report);
    }
}
