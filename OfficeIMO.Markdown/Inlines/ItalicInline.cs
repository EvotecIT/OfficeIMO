namespace OfficeIMO.Markdown;

/// <summary>
/// Italic emphasis inline.
/// </summary>
public sealed class ItalicInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates an italic inline with the given text.</summary>
    public ItalicInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => "_" + MarkdownEscaper.EscapeEmphasis(Text) + "_";
    internal string RenderHtml() => "<em>" + System.Net.WebUtility.HtmlEncode(Text) + "</em>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
