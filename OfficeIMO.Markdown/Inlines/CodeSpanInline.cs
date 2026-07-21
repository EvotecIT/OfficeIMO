namespace OfficeIMO.Markdown;

/// <summary>
/// Inline code span.
/// </summary>
public sealed class CodeSpanInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Code content.</summary>
    public string Text { get; }
    /// <summary>Source span for the code content token when parsed from markdown.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; internal set; }
    /// <summary>Creates an inline code span.</summary>
    public CodeSpanInline(string text) { Text = text ?? string.Empty; }

    internal void SetMarkdownSyntaxMetadataSpans(MarkdownSourceSpan? contentSourceSpan) {
        ContentSourceSpan = contentSourceSpan;
    }

    internal string RenderMarkdown() {
        return MarkdownFence.BuildSafeCodeSpan(Text);
    }
    internal string RenderHtml() => "<code>" + HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options) + "</code>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
