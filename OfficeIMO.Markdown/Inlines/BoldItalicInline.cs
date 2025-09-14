namespace OfficeIMO.Markdown;

/// <summary>
/// Combined bold+italic inline, rendered as ***text*** in Markdown and <strong><em>text</em></strong> in HTML.
/// </summary>
public sealed class BoldItalicInline {
    public string Text { get; }
    public BoldItalicInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => "***" + Text.Replace("***", "\\***") + "***";
    internal string RenderHtml() => "<strong><em>" + System.Net.WebUtility.HtmlEncode(Text) + "</em></strong>";
}

