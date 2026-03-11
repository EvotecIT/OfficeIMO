namespace OfficeIMO.Markdown;

internal interface IOwnedSyntaxChildrenMarkdownBlock {
    IReadOnlyList<MarkdownSyntaxNode> BuildOwnedSyntaxChildren();
}
