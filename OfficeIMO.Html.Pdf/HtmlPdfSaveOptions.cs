using MarkdownHtml = OfficeIMO.Markdown.Html;
using MarkdownPdf = OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using WordHtml = OfficeIMO.Word.Html;
using WordPdf = OfficeIMO.Word.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Options for the first-party HTML to PDF adapter.
/// </summary>
public sealed class HtmlPdfSaveOptions {
    /// <summary>
    /// Internal conversion path used before rendering through <see cref="PdfCore.PdfDocument"/>.
    /// Defaults to the semantic Markdown-backed path.
    /// </summary>
    public HtmlPdfProfile Profile { get; set; } = HtmlPdfProfile.Semantic;

    /// <summary>
    /// HTML to Markdown options used by the semantic profile.
    /// </summary>
    public MarkdownHtml.HtmlToMarkdownOptions? MarkdownHtmlOptions { get; set; }

    /// <summary>
    /// Markdown to PDF options used by the semantic profile.
    /// </summary>
    public MarkdownPdf.MarkdownPdfSaveOptions? MarkdownPdfOptions { get; set; }

    /// <summary>
    /// HTML to Word options used by the document profile.
    /// </summary>
    public WordHtml.HtmlToWordOptions? WordHtmlOptions { get; set; }

    /// <summary>
    /// Word to PDF options used by the document profile.
    /// </summary>
    public WordPdf.PdfSaveOptions? WordPdfOptions { get; set; }

    /// <summary>
    /// Shared conversion report populated from the selected internal HTML and PDF path.
    /// </summary>
    public PdfCore.PdfConversionReport ConversionReport { get; } = new PdfCore.PdfConversionReport();

    internal void ResetExportState() {
        ConversionReport.Clear();
    }
}
