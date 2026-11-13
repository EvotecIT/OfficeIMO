namespace OfficeIMO.Markdown;

/// <summary>
/// Text run that preserves semantic text while escaping Markdown block markers during serialization.
/// </summary>
internal sealed class LineStartEscapedTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IHtmlAppendableMarkdownInline {
    public LineStartEscapedTextRun(string text) {
        Text = text ?? string.Empty;
    }

    public string Text { get; }

    string IRenderableMarkdownInline.RenderMarkdown() => MarkdownEscaper.EscapeLineStartTextForWriter(Text);

    string IRenderableMarkdownInline.RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => HtmlTextEncoder.Append(builder, Text, HtmlRenderContext.Options);

    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
