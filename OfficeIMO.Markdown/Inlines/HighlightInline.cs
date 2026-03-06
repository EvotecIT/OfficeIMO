namespace OfficeIMO.Markdown;

/// <summary>
/// Highlighted inline text rendered as <c>==text==</c> in Markdown and <c>&lt;mark&gt;</c> in HTML.
/// </summary>
public sealed class HighlightInline {
    /// <summary>Text content.</summary>
    public string Text { get; }

    /// <summary>Creates a new highlighted inline.</summary>
    public HighlightInline(string text) {
        Text = text ?? string.Empty;
    }

    internal string RenderMarkdown() => "==" + MarkdownEscaper.EscapeHighlightText(Text) + "==";
    internal string RenderHtml() => "<mark>" + System.Net.WebUtility.HtmlEncode(Text) + "</mark>";
}
