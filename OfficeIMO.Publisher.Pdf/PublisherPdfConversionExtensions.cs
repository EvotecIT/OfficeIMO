using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Pdf;

/// <summary>Page-preserving PDF conversion of recovered Publisher publications.</summary>
public static class PublisherPdfConversionExtensions {
    /// <summary>Converts document pages in publication order, carrying source recovery losses into the PDF result.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(this PublisherDocument document, PdfCreateOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var report = new PdfConversionReport();
        PdfCreateOptions operation = options?.Clone() ?? new PdfCreateOptions();
        operation.ReportDiagnosticsTo(report, "OfficeIMO.Publisher.Pdf");
        PdfDocument pdf = PdfDocument.Create(operation);
        foreach (PublisherPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            pdf.Compose(builder => builder.Page(surface => surface.Size(page.Width, page.Height).Margin(0)
                .Canvas(canvas => canvas.Drawing(page.Drawing, 0, 0, page.Width, page.Height))));
        }
        return new PdfDocumentConversionResult(pdf, report).WithSourceConversionReport(document.ReadReport);
    }
    /// <summary>Converts to the PDF document model. Use ToPdfDocumentResult to inspect fidelity evidence.</summary>
    public static PdfDocument ToPdfDocument(this PublisherDocument document, PdfCreateOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Value;
    /// <summary>Converts document pages to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this PublisherDocument document, PdfCreateOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);
    /// <summary>Saves all document pages as PDF, returning source and output fidelity evidence.</summary>
    public static PdfSaveResult SaveAsPdf(this PublisherDocument document, string path, PdfCreateOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);
    /// <summary>Writes all document pages as PDF to a caller-owned stream without closing it.</summary>
    public static PdfSaveResult SaveAsPdf(this PublisherDocument document, Stream stream, PdfCreateOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);
}
