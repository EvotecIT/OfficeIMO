namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text produced by decoding HTML entities inside supported inline HTML wrappers.
/// Keeps the decoded text in the AST while re-encoding angle brackets during Markdown rendering
/// so literal tag text does not become active HTML on roundtrip.
/// </summary>
internal sealed class DecodedHtmlEntityTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    internal DecodedHtmlEntityTextRun(string text) {
        Text = text ?? string.Empty;
    }

    internal string Text { get; }

    string IRenderableMarkdownInline.RenderMarkdown() => MarkdownEscaper.EscapeLiteralText(Text);
    string IRenderableMarkdownInline.RenderHtml() => System.Net.WebUtility.HtmlEncode(Text);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
