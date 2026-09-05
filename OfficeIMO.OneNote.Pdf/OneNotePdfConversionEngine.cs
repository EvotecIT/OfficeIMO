using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNotePdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteSection section, OneNoteToPdfOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteToPdfOptions operation = (options ?? new OneNoteToPdfOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = section.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions);
    }

    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteNotebook notebook, OneNoteToPdfOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNoteToPdfOptions operation = (options ?? new OneNoteToPdfOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = notebook.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.MarkdownOptions);
    }

    private static PdfCore.PdfDocumentConversionResult ConvertProjection(
        OneNoteMarkdownConversionResult projection,
        MarkdownToPdfOptions pdfOptions) {
        if (pdfOptions.TextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.TextFallbacks |= PdfCore.PdfTextFallbackFeatures.MultilingualFonts;
        }

        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(pdfOptions);
        return result.WithSourceConversionReport(projection.Report);
    }
}
