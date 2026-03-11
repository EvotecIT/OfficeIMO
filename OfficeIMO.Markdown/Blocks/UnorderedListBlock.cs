namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <summary>Read-only AST-style view of list items.</summary>
    public IReadOnlyList<ListItem> ListItems => Items;
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

    IReadOnlyList<ListItem> IMarkdownListBlock.ListItems => ListItems;
    MarkdownSyntaxKind IMarkdownListBlock.ListSyntaxKind => MarkdownSyntaxKind.UnorderedList;
    string? IMarkdownListBlock.ListLiteral => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownListSyntax.BuildListBlockNode(this, span);
}
