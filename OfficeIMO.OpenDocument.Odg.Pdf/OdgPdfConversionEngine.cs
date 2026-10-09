using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odg.Pdf;

internal static class OdgPdfConversionEngine {
    internal static PdfDocumentConversionResult Convert(OdgDocument source, OdgToPdfOptions? options, CancellationToken token) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        token.ThrowIfCancellationRequested();
        OdgToPdfOptions operation = (options ?? new OdgToPdfOptions()).Snapshot();
        var pdfReport = new PdfConversionReport();
        PdfOptions pdfOptions = operation.PdfOptions ?? new PdfOptions();
        pdfOptions.ReportDiagnosticsTo(pdfReport, "OfficeIMO.OpenDocument.Odg.Pdf");
        var profile = new OfficeRenderingProfile("odg-pdf",
            textShapingProvider: pdfOptions.TextShapingProvider, textShapingLanguage: pdfOptions.Language);
        OdfConversionResult<IReadOnlyList<OfficeDrawing>> projected = source.ToDrawings(operation.LossPolicy, operation.ForPrint,
            operation.MaximumPages, token, operation.DateTimeFields, profile, PdfWriter.CreateDrawingTextMetrics(pdfOptions, token));
        PdfDocument pdf = PdfDocument.Create(pdfOptions);
        foreach (OfficeDrawing drawing in projected.Value) {
            token.ThrowIfCancellationRequested();
            pdf.Compose(document => document.Page(builder => {
                builder.Size(drawing.Width, drawing.Height).Margin(0);
                builder.Canvas(canvas => {
                    if (drawing.Elements.Count > 0) canvas.Drawing(drawing, 0, 0, drawing.Width, drawing.Height);
                });
            }));
        }
        token.ThrowIfCancellationRequested();
        return new PdfDocumentConversionResult(pdf, pdfReport).WithSourceConversionReport(projected.Report);
    }
}
