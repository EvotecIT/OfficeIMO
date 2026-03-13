using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Reader.Html;

/// <summary>
/// Options for HTML ingestion pipeline (HTML -> Markdown).
/// </summary>
public sealed class ReaderHtmlOptions {
    /// <summary>
    /// Options passed to HTML-to-Markdown conversion stage.
    /// </summary>
    public HtmlToMarkdownOptions? HtmlToMarkdownOptions { get; set; }
}
