namespace OfficeIMO.Markdown;

/// <summary>
/// Text run that preserves semantic text while escaping Markdown block markers during serialization.
/// </summary>
internal sealed class LineStartEscapedTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    public LineStartEscapedTextRun(string text) {
        Text = text ?? string.Empty;
    }

    public string Text { get; }

    string IRenderableMarkdownInline.RenderMarkdown() => MarkdownEscaper.EscapeTextAndLineStarts(Text);

    string IRenderableMarkdownInline.RenderHtml() => System.Net.WebUtility.HtmlEncode(Text);

    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
