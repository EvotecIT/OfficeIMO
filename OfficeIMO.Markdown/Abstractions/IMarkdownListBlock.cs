namespace OfficeIMO.Markdown;

internal interface IMarkdownListBlock : IMarkdownBlock, ISyntaxMarkdownBlock {
    IReadOnlyList<ListItem> ListItems { get; }
    MarkdownSyntaxKind ListSyntaxKind { get; }
    string? ListLiteral { get; }
}
