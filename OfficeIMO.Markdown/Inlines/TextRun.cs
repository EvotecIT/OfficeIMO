namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class TextRun {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a plain text run.</summary>
    public TextRun(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);
    internal string RenderHtml() => System.Net.WebUtility.HtmlEncode(Text);
}
