namespace OfficeIMO.Markdown;

public sealed class LinkInline {
    public string Text { get; }
    public string Url { get; }
    public string? Title { get; }
    public LinkInline(string text, string url, string? title) { Text = text ?? string.Empty; Url = url ?? string.Empty; Title = title; }
    internal string RenderMarkdown() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : " \"" + Title + "\"";
        return $"[{Text}]({Url}{title})";
    }
    internal string RenderHtml() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        return $"<a href=\"{System.Net.WebUtility.HtmlEncode(Url)}\"{title}>{System.Net.WebUtility.HtmlEncode(Text)}</a>";
    }
}

