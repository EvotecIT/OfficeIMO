namespace OfficeIMO.Markdown;

/// <summary>
/// List item content; supports plain and task (checklist) items.
/// </summary>
public sealed class ListItem {
    /// <summary>Inlines representing item content.</summary>
    public InlineSequence Content { get; }
    /// <summary>Additional paragraphs inside the list item (multi-paragraph list items).</summary>
    public List<InlineSequence> AdditionalParagraphs { get; } = new List<InlineSequence>();
    /// <summary>Nested block content inside the list item (e.g., nested ordered/unordered lists, code blocks).</summary>
    public List<IMarkdownBlock> Children { get; } = new List<IMarkdownBlock>();
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
    /// <summary>
    /// Creates a task (checklist) item using inline content.
    /// </summary>
    /// <param name="content">Inline content for the list item. When <c>null</c>, an empty sequence is used.</param>
    /// <param name="done">Whether the task should be marked as completed.</param>
    public static ListItem TaskInlines(InlineSequence content, bool done = false) => new ListItem(content ?? new InlineSequence(), true, done);

    internal IEnumerable<InlineSequence> Paragraphs() {
        yield return Content;
        for (int i = 0; i < AdditionalParagraphs.Count; i++) yield return AdditionalParagraphs[i];
    }

    internal string RenderMarkdown() {
        var parts = Paragraphs().Select(p => p.RenderMarkdown());
        return string.Join("\n\n", parts);
    }

    internal string RenderHtml() {
        string checkbox = IsTask ? "<input type=\"checkbox\" disabled" + (Checked ? " checked" : string.Empty) + "> " : string.Empty;
        if (AdditionalParagraphs.Count == 0 && Children.Count == 0) {
            return checkbox + Content.RenderHtml();
        }

        // Tight list behavior: when there is exactly one paragraph, keep it inline even if child blocks exist.
        if (AdditionalParagraphs.Count == 0) {
            var sbTight = new StringBuilder();
            sbTight.Append(checkbox).Append(Content.RenderHtml());
            for (int i = 0; i < Children.Count; i++) {
                if (Children[i] is IMarkdownBlock b) sbTight.Append(b.RenderHtml());
            }
            return sbTight.ToString();
        }

        // When multiple paragraphs exist, wrap paragraph content in <p> tags.
        var sb = new StringBuilder();
        bool first = true;
        foreach (var p in Paragraphs()) {
            sb.Append("<p>");
            if (first && IsTask) sb.Append(checkbox);
            sb.Append(p.RenderHtml());
            sb.Append("</p>");
            first = false;
        }

        for (int i = 0; i < Children.Count; i++) {
            if (Children[i] is IMarkdownBlock b) sb.Append(b.RenderHtml());
        }
        return sb.ToString();
    }
    internal string ToMarkdownListLine() {
        var indent = new string(' ', Level * 2);
        if (IsTask) return indent + "- [" + (Checked ? "x" : " ") + "] " + RenderMarkdown();
        return indent + "- " + RenderMarkdown();
    }
    internal string ToHtmlListItem() {
        return "<li>" + RenderHtml() + "</li>";
    }
}
