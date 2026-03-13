namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Options controlling HTML to Markdown conversion.
/// </summary>
public sealed class HtmlToMarkdownOptions {
    /// <summary>
    /// Optional base URI used to resolve relative links and image sources.
    /// </summary>
    public Uri? BaseUri { get; set; }

    /// <summary>
    /// When true, only the body contents are converted when a body element is present.
    /// </summary>
    public bool UseBodyContentsOnly { get; set; } = true;

    /// <summary>
    /// When true, script/style/noscript/template elements are ignored.
    /// </summary>
    public bool RemoveScriptsAndStyles { get; set; } = true;

    /// <summary>
    /// When true, unsupported block elements are emitted as raw HTML blocks instead of being dropped.
    /// </summary>
    public bool PreserveUnsupportedBlocks { get; set; } = true;

    /// <summary>
    /// When true, unsupported inline elements are emitted as raw HTML inside inline Markdown.
    /// </summary>
    public bool PreserveUnsupportedInlineHtml { get; set; } = true;
}
