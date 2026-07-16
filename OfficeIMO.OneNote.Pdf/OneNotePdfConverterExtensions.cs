using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Converts typed offline OneNote models to PDF through the shared Markdown model.</summary>
public static class OneNotePdfConverterExtensions {
    /// <summary>Converts a section to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).ToPdfDocument(PreparePdfOptions(pdfOptions));

    /// <summary>Converts a notebook to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).ToPdfDocument(PreparePdfOptions(pdfOptions));

    /// <summary>Converts a section to PDF bytes.</summary>
    public static byte[] ToPdf(
        this OneNoteSection section,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).ToPdf(PreparePdfOptions(pdfOptions));

    /// <summary>Converts a notebook to PDF bytes.</summary>
    public static byte[] ToPdf(
        this OneNoteNotebook notebook,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).ToPdf(PreparePdfOptions(pdfOptions));

    /// <summary>Saves a section as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(
        this OneNoteSection section,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsPdf(path, PreparePdfOptions(pdfOptions));

    /// <summary>Saves a notebook as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsPdf(path, PreparePdfOptions(pdfOptions));

    /// <summary>Writes a section as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(
        this OneNoteSection section,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsPdf(stream, PreparePdfOptions(pdfOptions));

    /// <summary>Writes a notebook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfDocumentConversionResult SaveAsPdf(
        this OneNoteNotebook notebook,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsPdf(stream, PreparePdfOptions(pdfOptions));

    /// <summary>Asynchronously saves a section as PDF.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this OneNoteSection section,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsPdfAsync(path, PreparePdfOptions(pdfOptions), cancellationToken);

    /// <summary>Asynchronously saves a notebook as PDF.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this OneNoteNotebook notebook,
        string path,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsPdfAsync(path, PreparePdfOptions(pdfOptions), cancellationToken);

    /// <summary>Asynchronously writes a section as PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this OneNoteSection section,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) =>
        section.ToMarkdownDocument(projectionOptions).SaveAsPdfAsync(stream, PreparePdfOptions(pdfOptions), cancellationToken);

    /// <summary>Asynchronously writes a notebook as PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(
        this OneNoteNotebook notebook,
        Stream stream,
        OneNoteMarkdownOptions? projectionOptions = null,
        MarkdownPdfSaveOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) =>
        notebook.ToMarkdownDocument(projectionOptions).SaveAsPdfAsync(stream, PreparePdfOptions(pdfOptions), cancellationToken);

    private static MarkdownPdfSaveOptions PreparePdfOptions(MarkdownPdfSaveOptions? options) {
        MarkdownPdfSaveOptions prepared = options?.Clone() ?? new MarkdownPdfSaveOptions();
        if (prepared.TextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            prepared.TextFallbacks |= PdfCore.PdfTextFallbackFeatures.MultilingualFonts;
        }
        return prepared;
    }
}
