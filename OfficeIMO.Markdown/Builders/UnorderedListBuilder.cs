namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for unordered lists.
/// </summary>
public sealed class UnorderedListBuilder {
    private readonly UnorderedListBlock _ul = new UnorderedListBlock();
    /// <summary>Adds a plain text item.</summary>
    public UnorderedListBuilder Item(string text) { _ul.Items.Add(ListItem.Text(text)); return this; }
    /// <summary>Adds a hyperlink item.</summary>
    public UnorderedListBuilder ItemLink(string text, string url, string? title = null) { _ul.Items.Add(ListItem.Link(text, url, title)); return this; }
    /// <summary>Adds a task (checklist) item.</summary>
    public UnorderedListBuilder ItemTask(string text, bool done = false) { _ul.Items.Add(ListItem.Task(text, done)); return this; }
    internal UnorderedListBlock Build() => _ul;
}
