namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered (numbered) list.
/// </summary>
public sealed class OrderedListBlock : IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>Items within the ordered list.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <summary>Read-only AST-style view of list items.</summary>
    public IReadOnlyList<ListItem> ListItems => Items;
    /// <summary>Starting number (default 1).</summary>
    public int Start { get; set; } = 1;

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() =>
        MarkdownListRendering.RenderMarkdown(
            Items,
            (item, topLevelIndex) => {
                string baseMarker = item.Level == 0
                    ? (Start + topLevelIndex).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". "
                    : "1. ";
                return item.IsTask
                    ? baseMarker + "[" + (item.Checked ? "x" : " ") + "] "
                    : baseMarker;
            });

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() =>
        MarkdownListRendering.RenderHtml(
            "ol",
            Items,
            _ => Start != 1 ? " start=\"" + Start + "\"" : string.Empty);

    IReadOnlyList<ListItem> IMarkdownListBlock.ListItems => ListItems;
    MarkdownSyntaxKind IMarkdownListBlock.ListSyntaxKind => MarkdownSyntaxKind.OrderedList;
    string? IMarkdownListBlock.ListLiteral => Start.ToString(System.Globalization.CultureInfo.InvariantCulture);
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownListSyntax.BuildListBlockNode(this, span);
}
