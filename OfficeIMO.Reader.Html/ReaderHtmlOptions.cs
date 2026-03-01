using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Reader.Html;

/// <summary>
/// Options for HTML ingestion pipeline (HTML -> Word -> Markdown).
/// </summary>
public sealed class ReaderHtmlOptions {
    /// <summary>
    /// Options passed to HTML-to-Word conversion stage.
    /// </summary>
    public HtmlToWordOptions? HtmlToWordOptions { get; set; }

    /// <summary>
    /// Options passed to Word-to-Markdown conversion stage.
    /// </summary>
    public WordToMarkdownOptions? MarkdownOptions { get; set; }
}
