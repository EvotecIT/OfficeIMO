using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNotePdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteSection section, OneNoteToPdfOptions? options, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteToPdfOptions operation = (options ?? new OneNoteToPdfOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = section.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions, cancellationToken);
    }

    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteNotebook notebook, OneNoteToPdfOptions? options, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteToPdfOptions operation = (options ?? new OneNoteToPdfOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = notebook.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions, cancellationToken);
    }

    private static PdfCore.PdfDocumentConversionResult ConvertProjection(
        OneNoteMarkdownConversionResult projection,
        MarkdownToPdfOptions pdfOptions, System.Threading.CancellationToken cancellationToken) {
        if (pdfOptions.TextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.TextFallbacks |= PdfCore.PdfTextFallbackFeatures.MultilingualFonts;
        }

        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(pdfOptions, cancellationToken);
        return result.WithSourceConversionReport(projection.Report);
    }
}
