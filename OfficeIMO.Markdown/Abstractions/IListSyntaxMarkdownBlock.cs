namespace OfficeIMO.Markdown;

internal interface IListSyntaxMarkdownBlock {
    IReadOnlyList<ListItem> ListItems { get; }
    MarkdownSyntaxKind ListSyntaxKind { get; }
    string? ListLiteral { get; }
}
