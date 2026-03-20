namespace OfficeIMO.Markdown;

/// <summary>
/// Horizontal rule (thematic break). Rendered as --- in Markdown and <hr /> in HTML.
/// </summary>
public sealed class HorizontalRuleBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    string IMarkdownBlock.RenderMarkdown() => "---";
    string IMarkdownBlock.RenderHtml() => "<hr />";
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.HorizontalRule, span, "---", associatedObject: this);
}
