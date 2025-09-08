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
    /// <summary>Adds multiple items from a sequence of values using <c>ToString()</c>.</summary>
    public UnorderedListBuilder Items(System.Collections.Generic.IEnumerable<string> items) { foreach (var i in items) _ul.Items.Add(ListItem.Text(i)); return this; }
    /// <summary>Adds multiple items using a selector to format each element.</summary>
    public UnorderedListBuilder Items<T>(System.Collections.Generic.IEnumerable<T> items, System.Func<T, string>? selector = null) {
        selector ??= (x => x?.ToString() ?? string.Empty);
        foreach (var i in items) _ul.Items.Add(ListItem.Text(selector(i)));
        return this;
    }
    internal UnorderedListBlock Build() => _ul;
}
