using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Pdf;

namespace OfficeIMO.Notebook.Pdf;

internal static class NotebookPdfConversionEngine {
    internal static PdfDocumentConversionResult Convert(
        NotebookDocument notebook, NotebookPdfExportOptions? options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        NotebookPdfExportOptions operation = (options ?? new NotebookPdfExportOptions()).Clone();
        MarkdownPdfExportOptions pdfOptions = operation.MarkdownOptions;
        PdfResourcePolicy resourcePolicy = pdfOptions.ResourcePolicy.Clone();
        resourcePolicy.AllowLocalFileAccess = false;
        resourcePolicy.AllowRemoteResourceResolution = false;
        pdfOptions.ResourcePolicy = resourcePolicy;
        pdfOptions.BaseDirectory = null;
        pdfOptions.RemoteImageResolver = null;
        pdfOptions.Title = operation.Title ?? pdfOptions.Title;
        pdfOptions.Author = operation.Author ?? pdfOptions.Author;
        pdfOptions.UseFrontMatterMetadata = false;
        pdfOptions.UseFrontMatterTheme = false;
        pdfOptions.PaginateFigures = true;
        if (pdfOptions.TextFallbacks != PdfTextFallbackFeatures.None) {
            pdfOptions.TextFallbacks |= PdfTextFallbackFeatures.MultilingualFonts;
        }

        using MarkdownDocument projection = notebook.ToMarkdown(operation.CreateProjectionOptions(), out OfficeConversionReport projectionReport, cancellationToken);
        PdfDocumentConversionResult result = projection.ConvertToPdfWithEvidence(pdfOptions, cancellationToken);
        return result.WithSourceConversionReport(projectionReport);
    }
}
