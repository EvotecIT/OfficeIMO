namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class TextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a plain text run.</summary>
    public TextRun(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);
    internal string RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
