namespace OfficeIMO.Markdown;

internal interface ISyntaxMarkdownBlock {
    MarkdownSyntaxNode BuildSyntaxNode(MarkdownSourceSpan? span);
}
