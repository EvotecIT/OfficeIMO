using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using System.Linq;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNotePdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteSection section, OneNotePdfSaveOptions? options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNotePdfSaveOptions operation = (options ?? new OneNotePdfSaveOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = section.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.PdfOptions);
    }

    internal static PdfCore.PdfDocumentConversionResult Convert(OneNoteNotebook notebook, OneNotePdfSaveOptions? options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        OneNotePdfSaveOptions operation = (options ?? new OneNotePdfSaveOptions()).CloneForConversion();
        OneNoteMarkdownConversionResult projection = notebook.ToMarkdownDocumentResult(operation.ProjectionOptions);
        return ConvertProjection(projection, operation.PdfOptions);
    }

    private static PdfCore.PdfDocumentConversionResult ConvertProjection(
        OneNoteMarkdownConversionResult projection,
        MarkdownPdfSaveOptions pdfOptions) {
        if (pdfOptions.TextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdfOptions.TextFallbacks |= PdfCore.PdfTextFallbackFeatures.MultilingualFonts;
        }

        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(pdfOptions);
        return result.WithAdditionalWarnings(projection.Diagnostics.Select(ToPdfWarning));
    }

    private static PdfCore.PdfConversionWarning ToPdfWarning(OneNoteMarkdownDiagnostic diagnostic) =>
        new PdfCore.PdfConversionWarning(
            "OfficeIMO.OneNote.Pdf",
            diagnostic.Code,
            diagnostic.Source,
            diagnostic.Message,
            diagnostic.Severity == OneNoteDiagnosticSeverity.Error
                ? PdfCore.PdfConversionWarningSeverity.Error
                : diagnostic.Severity == OneNoteDiagnosticSeverity.Information
                    ? PdfCore.PdfConversionWarningSeverity.Information
                    : PdfCore.PdfConversionWarningSeverity.Warning);
}
