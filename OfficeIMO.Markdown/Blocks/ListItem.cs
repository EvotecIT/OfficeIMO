namespace OfficeIMO.Markdown;

public sealed class ListItem {
    public InlineSequence Content { get; }
    public ListItem(InlineSequence content) { Content = content; }
    public static ListItem Text(string text) => new ListItem(new InlineSequence().Text(text));
    public static ListItem Link(string text, string url, string? title = null) => new ListItem(new InlineSequence().Link(text, url, title));
    internal string RenderMarkdown() => Content.RenderMarkdown();
    internal string RenderHtml() => Content.RenderHtml();
}

