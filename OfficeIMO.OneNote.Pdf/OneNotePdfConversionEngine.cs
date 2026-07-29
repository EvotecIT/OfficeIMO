using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNotePdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteSection section, OneNotePdfSaveOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNotePdfSaveOptions operation = (options ?? new OneNotePdfSaveOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = section.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions);
    }

    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteNotebook notebook, OneNotePdfSaveOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNotePdfSaveOptions operation = (options ?? new OneNotePdfSaveOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = notebook.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions);
    }

    private static PdfCore.PdfDocumentConversionResult ConvertProjection(
        OneNoteMarkdownConversionResult projection,
        MarkdownPdfSaveOptions pdfOptions) {
        if (pdfOptions.TextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.TextFallbacks |= PdfCore.PdfTextFallbackFeatures.MultilingualFonts;
        }

        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(pdfOptions);
        return result.WithSourceConversionReport(projection.Report);
    }
}
