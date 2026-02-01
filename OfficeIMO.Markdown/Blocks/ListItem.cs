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
    /// <summary>Indentation level (0 = top-level). Used for nested lists.</summary>
    public int Level { get; set; }

    /// <summary>Creates a plain list item.</summary>
    public ListItem(InlineSequence content) { Content = content; }
    private ListItem(InlineSequence content, bool isTask, bool isChecked) { Content = content; IsTask = isTask; Checked = isChecked; }

    /// <summary>Creates a plain text item.</summary>
    public static ListItem Text(string text) => new ListItem(new InlineSequence().Text(text));
    /// <summary>Creates a link item.</summary>
    public static ListItem Link(string text, string url, string? title = null) => new ListItem(new InlineSequence().Link(text, url, title));
    /// <summary>Creates a task (checklist) item.</summary>
    public static ListItem Task(string text, bool done = false) => new ListItem(new InlineSequence().Text(text), true, done);
    /// <summary>Creates a task (checklist) item with inline markup parsed.</summary>
    public static ListItem Task(InlineSequence content, bool done = false) => new ListItem(content ?? new InlineSequence(), true, done);

    internal string RenderMarkdown() => Content.RenderMarkdown();
    internal string RenderHtml() => Content.RenderHtml();
    internal string ToMarkdownListLine() {
        var indent = new string(' ', Level * 2);
        if (IsTask) return indent + "- [" + (Checked ? "x" : " ") + "] " + RenderMarkdown();
        return indent + "- " + RenderMarkdown();
    }
    internal string ToHtmlListItem() {
        if (IsTask) {
            string checkbox = "<input type=\"checkbox\" disabled" + (Checked ? " checked" : string.Empty) + "> ";
            return "<li>" + checkbox + RenderHtml() + "</li>";
        }
        return "<li>" + RenderHtml() + "</li>";
    }
}
