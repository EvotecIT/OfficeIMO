using OfficeIMO.Pdf;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Pdf;

/// <summary>Page-preserving PDF conversion of recovered Publisher publications.</summary>
public static class PublisherPdfConversionExtensions {
    /// <summary>Converts document pages in publication order, carrying source recovery losses into the PDF result.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(this PublisherDocument document, PdfOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var report = new PdfConversionReport();
        PdfOptions operation = options?.Clone() ?? new PdfOptions();
        var drawings = new System.Collections.Generic.List<OfficeDrawing>(document.Pages.Count);
        foreach (PublisherPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            drawings.Add(page.CreateDrawing());
        }
        var fonts = new OfficeFontFaceCollection();
        foreach (OfficeDrawing drawing in drawings) fonts.AddRange(drawing.Fonts);
        if (fonts.Faces.Count != 0)
            operation.UseRenderingProfile(new OfficeRenderingProfile("Publisher", fonts), OfficeRenderingProfileApplyMode.Overlay);
        operation.ReportDiagnosticsTo(report, "OfficeIMO.Publisher.Pdf");
        PdfDocument pdf = PdfDocument.Create(operation);
        foreach (OfficeDrawing drawing in drawings) {
            cancellationToken.ThrowIfCancellationRequested();
            pdf.Compose(builder => builder.Page(surface => surface.Size(drawing.Width, drawing.Height).Margin(0)
                .Canvas(canvas => canvas.Drawing(drawing, 0, 0, drawing.Width, drawing.Height))));
        }
        return new PdfDocumentConversionResult(pdf, report).WithSourceConversionReport(document.ReadReport);
    }
    /// <summary>Converts to the PDF document model. Use ToPdfDocumentResult to inspect fidelity evidence.</summary>
    public static PdfDocument ToPdfDocument(this PublisherDocument document, PdfOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Value;
    /// <summary>Converts document pages to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this PublisherDocument document, PdfOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);
    /// <summary>Saves all document pages as PDF, returning source and output fidelity evidence.</summary>
    public static PdfSaveResult SaveAsPdf(this PublisherDocument document, string path, PdfOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);
    /// <summary>Writes all document pages as PDF to a caller-owned stream without closing it.</summary>
    public static PdfSaveResult SaveAsPdf(this PublisherDocument document, Stream stream, PdfOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);
}
