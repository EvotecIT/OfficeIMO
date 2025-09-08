namespace OfficeIMO.Markdown;

public sealed class UnorderedListBuilder {
    private readonly UnorderedListBlock _ul = new UnorderedListBlock();
    public UnorderedListBuilder Item(string text) { _ul.Items.Add(ListItem.Text(text)); return this; }
    public UnorderedListBuilder ItemLink(string text, string url, string? title = null) { _ul.Items.Add(ListItem.Link(text, url, title)); return this; }
    internal UnorderedListBlock Build() => _ul;
}

