namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote reference inline, e.g., [^1].
/// </summary>
public sealed class FootnoteRefInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IContextualHtmlMarkdownInline {
    /// <summary>Reference label pointing to a footnote definition.</summary>
    public string Label { get; }
    /// <summary>Create a footnote reference inline.</summary>
    public FootnoteRefInline(string label) { Label = label ?? string.Empty; }
    internal string RenderMarkdown() => $"[^{Label}]";
    internal string RenderHtml() => $"<sup id=\"fnref:{System.Net.WebUtility.HtmlEncode(Label)}\"><a href=\"#fn:{System.Net.WebUtility.HtmlEncode(Label)}\">{System.Net.WebUtility.HtmlEncode(Label)}</a></sup>";
    string IContextualHtmlMarkdownInline.RenderHtml(HtmlOptions options) {
        if (options?.GitHubFootnoteHtml != true) {
            return RenderHtml();
        }

        var state = HtmlRenderContext.Footnotes;
        if (state == null || !state.IsDefined(Label)) {
            return System.Net.WebUtility.HtmlEncode(RenderMarkdown());
        }

        var reference = state.RegisterReference(Label);
        return "<sup class=\"footnote-ref\"><a href=\"#fn-"
               + System.Net.WebUtility.HtmlEncode(reference.EscapedLabel)
               + "\" id=\""
               + System.Net.WebUtility.HtmlEncode(reference.ReferenceId)
               + "\" data-footnote-ref>"
               + reference.Number.ToString()
               + "</a></sup>";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Label);
}
