namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() =>
        MarkdownListRendering.RenderMarkdown(
            Items,
            (item, _) => item.IsTask
                ? "- [" + (item.Checked ? "x" : " ") + "] "
                : "- ");
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() =>
        MarkdownListRendering.RenderHtml("ul", Items, _ => string.Empty);

    IReadOnlyList<ListItem> IListSyntaxMarkdownBlock.ListItems => Items;
    MarkdownSyntaxKind IListSyntaxMarkdownBlock.ListSyntaxKind => MarkdownSyntaxKind.UnorderedList;
    string? IListSyntaxMarkdownBlock.ListLiteral => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildListBlock(this, span);
}
