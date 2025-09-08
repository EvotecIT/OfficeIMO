namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience wrapper for an unordered list (object-model style).
/// </summary>
public sealed class UnorderedList : IMarkdownBlock {
    private readonly UnorderedListBlock _ul = new UnorderedListBlock();
    /// <summary>Adds an item to the list.</summary>
    public void Add(ListItem item) => _ul.Items.Add(item);
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_ul).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_ul).RenderHtml();
}
