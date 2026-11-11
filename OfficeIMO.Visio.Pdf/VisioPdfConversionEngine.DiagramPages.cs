using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static partial class VisioPdfConversionEngine {
    private static PdfDocumentConversionResult ConvertDiagramPages(VisioDocument source, VisioPdfExportOptions options, CancellationToken token) {
        VisioDrawingOptions drawingOptions = options.DrawingOptions?.Clone() ?? new VisioDrawingOptions();
        if (options.SourceName != null) drawingOptions.SourceName = options.SourceName;
        var report = new PdfConversionReport();
        PdfCreateOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCreateOptions();
        pdfOptions.ReportDiagnosticsTo(report, "OfficeIMO.Visio.Pdf");
        var profile = new OfficeRenderingProfile(
            "visio-diagram-pages",
            fonts: drawingOptions.Fonts,
            textShapingProvider: drawingOptions.TextShapingProvider,
            textShapingLanguage: drawingOptions.TextShapingLanguage);
        pdfOptions.UseRenderingProfile(profile, OfficeRenderingProfileApplyMode.Overlay);
        drawingOptions.LayoutMetrics = PdfWriter.CreateDrawingTextMetrics(pdfOptions, token);
        OfficeConversionResult<IReadOnlyList<OfficeDrawing>, VisioDrawingConversionReport> projected = source.ToDrawings(drawingOptions, token);
        PdfDocument pdf = PdfDocument.CreateForImportedLayout(pdfOptions);
        for (int index = 0; index < projected.Value.Count; index++) {
            token.ThrowIfCancellationRequested();
            OfficeDrawing drawing = projected.Value[index];
            string destination = VisioDrawingProjection.PageDestinationName(index + 1);
            pdf.Compose(document => document.Page(page => {
                page.Size(global::OfficeIMO.OfficeLength.Points(drawing.Width), global::OfficeIMO.OfficeLength.Points(drawing.Height)).Margin(global::OfficeIMO.OfficeLength.Points(0));
                page.Canvas(canvas => {
                    // Page hyperlinks in the scenes target these destinations.
                    canvas.NamedDestination(destination, global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0));
                    if (drawing.Elements.Count > 0) canvas.Drawing(drawing, global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(drawing.Width), global::OfficeIMO.OfficeLength.Points(drawing.Height));
                });
            }));
        }
        token.ThrowIfCancellationRequested();
        return new PdfDocumentConversionResult(pdf, report).WithSourceConversionReport(projected.Report);
    }
}
