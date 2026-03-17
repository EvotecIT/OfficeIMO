namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class TextRun : IMarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a plain text run.</summary>
    public TextRun(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => MarkdownEscaper.EscapeLiteralText(Text);
    internal string RenderHtml() => System.Net.WebUtility.HtmlEncode(Text);
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
