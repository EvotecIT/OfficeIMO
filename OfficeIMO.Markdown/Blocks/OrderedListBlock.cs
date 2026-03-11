namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered (numbered) list.
/// </summary>
public sealed class OrderedListBlock : IMarkdownListBlock, ISyntaxMarkdownBlock {
    /// <summary>Items within the ordered list.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
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

    IReadOnlyList<ListItem> IListSyntaxMarkdownBlock.ListItems => Items;
    MarkdownSyntaxKind IListSyntaxMarkdownBlock.ListSyntaxKind => MarkdownSyntaxKind.OrderedList;
    string? IListSyntaxMarkdownBlock.ListLiteral => Start.ToString(System.Globalization.CultureInfo.InvariantCulture);
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildListBlock(this, span);
}
