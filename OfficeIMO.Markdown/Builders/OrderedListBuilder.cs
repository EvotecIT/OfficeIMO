namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for ordered lists.
/// </summary>
public sealed class OrderedListBuilder {
    private readonly OrderedListBlock _ol = new OrderedListBlock();
    /// <summary>Sets the starting number (1-based).</summary>
    public OrderedListBuilder StartAt(int start) { _ol.Start = start < 1 ? 1 : start; return this; }
    /// <summary>Adds a plain text item.</summary>
    public OrderedListBuilder Item(string text) { _ol.Items.Add(ListItem.Text(text)); return this; }
    /// <summary>Adds a hyperlink item.</summary>
    public OrderedListBuilder ItemLink(string text, string url, string? title = null) { _ol.Items.Add(ListItem.Link(text, url, title)); return this; }
    internal OrderedListBlock Build() => _ol;
}
