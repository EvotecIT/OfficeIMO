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
    /// Creates the default semantic HTML to PDF profile backed by OfficeIMO.Markdown.Html and OfficeIMO.Markdown.Pdf.
    /// </summary>
    public static HtmlPdfSaveOptions CreateSemanticProfile() => new HtmlPdfSaveOptions {
        Profile = HtmlPdfProfile.Semantic,
        MarkdownHtmlOptions = MarkdownHtml.HtmlToMarkdownOptions.CreateOfficeIMOProfile(),
        MarkdownPdfOptions = new MarkdownPdf.MarkdownPdfSaveOptions()
    };

    /// <summary>
    /// Creates a document-oriented HTML to PDF profile backed by OfficeIMO.Word.Html and OfficeIMO.Word.Pdf.
    /// This path is the preferred adapter profile for practical print HTML with CSS, links, tables, images, and page-break hints.
    /// </summary>
    public static HtmlPdfSaveOptions CreateDocumentProfile() => new HtmlPdfSaveOptions {
        Profile = HtmlPdfProfile.Document,
        WordHtmlOptions = WordHtml.HtmlToWordOptions.CreateOfficeIMOProfile(),
        WordPdfOptions = new WordPdf.PdfSaveOptions()
    };

    /// <summary>
    /// Creates a document-oriented profile that allows stylesheet links declared by trusted HTML documents.
    /// </summary>
    public static HtmlPdfSaveOptions CreateTrustedDocumentProfile() => new HtmlPdfSaveOptions {
        Profile = HtmlPdfProfile.Document,
        WordHtmlOptions = WordHtml.HtmlToWordOptions.CreateTrustedDocumentProfile(),
        WordPdfOptions = new WordPdf.PdfSaveOptions()
    };

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

    /// <summary>
    /// Returns a source-neutral snapshot of the active HTML resource policy for manifest, wrapper, and diagnostics callers.
    /// </summary>
    public HtmlPdfResourcePolicySummary GetResourcePolicySummary() => HtmlPdfResourcePolicySummary.From(this);

    internal void ResetExportState() {
        ConversionReport.Clear();
        ConversionReport.ClearLinkedReports();
    }
}
