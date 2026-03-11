namespace OfficeIMO.Markdown;

/// <summary>
/// Horizontal rule (thematic break). Rendered as --- in Markdown and <hr /> in HTML.
/// </summary>
public sealed class HorizontalRuleBlock : IMarkdownBlock, ISyntaxMarkdownBlock {
    string IMarkdownBlock.RenderMarkdown() => "---";
    string IMarkdownBlock.RenderHtml() => "<hr />";
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildHorizontalRuleBlock(span);
}
