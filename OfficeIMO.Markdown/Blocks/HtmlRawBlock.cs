namespace OfficeIMO.Markdown;

/// <summary>
/// Raw HTML block passthrough.
/// </summary>
public sealed class HtmlRawBlock : IMarkdownBlock {
    public string Html { get; }
    public HtmlRawBlock(string html) { Html = html ?? string.Empty; }
    string IMarkdownBlock.RenderMarkdown() => Html;
    string IMarkdownBlock.RenderHtml() => Html;
}

