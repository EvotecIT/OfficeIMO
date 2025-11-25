namespace OfficeIMO.Markdown;

/// <summary>
/// Strikethrough inline (rendered as ~~text~~ in Markdown and as a deleted span in HTML).
/// </summary>
public sealed class StrikethroughInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a new strikethrough inline.</summary>
    public StrikethroughInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => $"~~{MarkdownEscaper.EscapeEmphasis(Text)}~~";
    internal string RenderHtml() => $"<del>{System.Net.WebUtility.HtmlEncode(Text)}</del>";
}
