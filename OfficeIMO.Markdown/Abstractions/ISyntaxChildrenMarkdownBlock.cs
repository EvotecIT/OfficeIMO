namespace OfficeIMO.Markdown;

internal interface ISyntaxChildrenMarkdownBlock {
    IReadOnlyList<MarkdownSyntaxNode>? ProvidedSyntaxChildren { get; }
}
