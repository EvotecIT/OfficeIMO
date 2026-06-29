namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : MarkdownBlock, IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <summary>Read-only AST-style view of list items.</summary>
    public IReadOnlyList<ListItem> ListItems => Items;
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() =>
        MarkdownListRendering.RenderMarkdown(
            Attributes,
            Items,
            (item, _) => {
                char marker = GetRenderMarker(item);
                return item.IsTask
                    ? marker + " [" + (item.Checked ? "x" : " ") + "] "
                    : marker + " ";
            });
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() =>
        MarkdownListRendering.RenderHtml("ul", Items, Attributes, _ => string.Empty);

    IReadOnlyList<ListItem> IMarkdownListBlock.ListItems => ListItems;
    MarkdownSyntaxKind IMarkdownListBlock.ListSyntaxKind => MarkdownSyntaxKind.UnorderedList;
    string? IMarkdownListBlock.ListLiteral => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownListSyntax.BuildListBlockNode(this, span);

    private static char GetRenderMarker(ListItem item) {
        if (item?.MarkerText is { Length: 1 } markerText &&
            (markerText[0] == '-' || markerText[0] == '*' || markerText[0] == '+')) {
            return markerText[0];
        }

        return MarkdownRenderContext.Options?.UnorderedListMarker ?? '-';
    }
}
