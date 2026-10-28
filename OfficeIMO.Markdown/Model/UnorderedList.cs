namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience wrapper for an unordered list (object-model style).
/// </summary>
public sealed class MarkdownUnorderedList : IMarkdownBlock {
    private readonly UnorderedListBlock _ul = new UnorderedListBlock();
    /// <summary>Adds an item to the list.</summary>
    /// <returns>The attached item.</returns>
    public ListItem Add(ListItem item) {
        if (item == null) throw new ArgumentNullException(nameof(item));
        _ul.Items.Add(item);
        return item;
    }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_ul).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_ul).RenderHtml();
}
