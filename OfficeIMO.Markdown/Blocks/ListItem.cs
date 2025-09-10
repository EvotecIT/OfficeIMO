namespace OfficeIMO.Markdown;

/// <summary>
/// List item content; supports plain and task (checklist) items.
/// </summary>
public sealed class ListItem {
    /// <summary>Inlines representing item content.</summary>
    public InlineSequence Content { get; }
    /// <summary>True when rendered as a task item (<c>- [ ]</c> or <c>- [x]</c>).</summary>
    public bool IsTask { get; }
    /// <summary>Whether the task is checked.</summary>
    public bool Checked { get; }

    /// <summary>Creates a plain list item.</summary>
    public ListItem(InlineSequence content) { Content = content; }
    private ListItem(InlineSequence content, bool isTask, bool isChecked) { Content = content; IsTask = isTask; Checked = isChecked; }

    /// <summary>Creates a plain text item.</summary>
    public static ListItem Text(string text) => new ListItem(new InlineSequence().Text(text));
    /// <summary>Creates a link item.</summary>
    public static ListItem Link(string text, string url, string? title = null) => new ListItem(new InlineSequence().Link(text, url, title));
    /// <summary>Creates a task (checklist) item.</summary>
    public static ListItem Task(string text, bool done = false) => new ListItem(new InlineSequence().Text(text), true, done);

    internal string RenderMarkdown() => Content.RenderMarkdown();
    internal string RenderHtml() => Content.RenderHtml();
    internal string ToMarkdownListLine() {
        if (IsTask) return "- [" + (Checked ? "x" : " ") + "] " + RenderMarkdown();
        return "- " + RenderMarkdown();
    }
    internal string ToHtmlListItem() {
        if (IsTask) {
            string checkbox = "<input type=\"checkbox\" disabled" + (Checked ? " checked" : string.Empty) + "> ";
            return "<li>" + checkbox + RenderHtml() + "</li>";
        }
        return "<li>" + RenderHtml() + "</li>";
    }
}
