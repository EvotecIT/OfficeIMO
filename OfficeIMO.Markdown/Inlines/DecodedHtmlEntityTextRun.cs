namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text produced by decoding HTML entities inside supported inline HTML wrappers.
/// Keeps the decoded text in the AST while re-encoding angle brackets during Markdown rendering
/// so literal tag text does not become active HTML on roundtrip.
/// </summary>
internal sealed class DecodedHtmlEntityTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline, IHtmlAppendableMarkdownInline {
    internal DecodedHtmlEntityTextRun(string text) {
        Text = text ?? string.Empty;
    }

    public string Text { get; }
    internal string? SourceText { get; private set; }
    internal MarkdownSourceSpan? SourceTextSourceSpan { get; private set; }

    internal void SetMarkdownSyntaxMetadataSpans(string? sourceText, MarkdownSourceSpan? sourceTextSourceSpan) {
        SourceText = sourceText;
        SourceTextSourceSpan = sourceTextSourceSpan;
    }

    string IRenderableMarkdownInline.RenderMarkdown() => MarkdownEscaper.EscapeLiteralTextForWriter(Text);
    string IRenderableMarkdownInline.RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => HtmlTextEncoder.Append(builder, Text, HtmlRenderContext.Options);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
