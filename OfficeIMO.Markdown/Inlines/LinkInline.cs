namespace OfficeIMO.Markdown;

/// <summary>
/// Hyperlink inline.
/// </summary>
public sealed class LinkInline {
    /// <summary>Link text.</summary>
    public string Text { get; }
    /// <summary>Destination URL.</summary>
    public string Url { get; }
    /// <summary>Optional title shown as a tooltip in HTML.</summary>
    public string? Title { get; }
    /// <summary>Creates a hyperlink inline.</summary>
    public LinkInline(string text, string url, string? title) { Text = text ?? string.Empty; Url = url ?? string.Empty; Title = title; }
    internal string RenderMarkdown() {
        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        return $"[{MarkdownEscaper.EscapeLinkText(Text)}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
    }
    internal string RenderHtml() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        return $"<a href=\"{System.Net.WebUtility.HtmlEncode(Url)}\"{title}>{System.Net.WebUtility.HtmlEncode(Text)}</a>";
    }
}
