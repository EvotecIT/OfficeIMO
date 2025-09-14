namespace OfficeIMO.Markdown;

/// <summary>
/// Horizontal rule (thematic break). Rendered as --- in Markdown and <hr /> in HTML.
/// </summary>
public sealed class HorizontalRuleBlock : IMarkdownBlock {
    string IMarkdownBlock.RenderMarkdown() => "---";
    string IMarkdownBlock.RenderHtml() => "<hr />";
}

