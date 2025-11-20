namespace OfficeIMO.Markdown;

/// <summary>
/// Combined bold+italic inline, rendered as ***text*** in Markdown and <strong><em>text</em></strong> in HTML.
/// </summary>
public sealed class BoldItalicInline {
    /// <summary>Content inside the emphasis.</summary>
    public string Text { get; }
    /// <summary>Create a bold+italic inline.</summary>
    public BoldItalicInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => "***" + MarkdownEscaper.EscapeEmphasis(Text) + "***";
    internal string RenderHtml() => "<strong><em>" + System.Net.WebUtility.HtmlEncode(Text) + "</em></strong>";
}
