namespace OfficeIMO.Markdown;

/// <summary>
/// Raw HTML block passthrough.
/// </summary>
public sealed class HtmlRawBlock : IMarkdownBlock {
    /// <summary>Raw HTML content to emit.</summary>
    public string Html { get; }
    /// <summary>Create a new raw HTML block.</summary>
    /// <param name="html">HTML fragment.</param>
    public HtmlRawBlock(string html) { Html = html ?? string.Empty; }
    string IMarkdownBlock.RenderMarkdown() => Html;
    string IMarkdownBlock.RenderHtml() => Html;
}
