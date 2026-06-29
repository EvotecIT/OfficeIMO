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
    internal string RenderHtml() {
        var options = HtmlRenderContext.Options;
        var encodedLabel = HtmlTextEncoder.Encode(Label, options);
        return $"<sup id=\"fnref:{encodedLabel}\"><a href=\"#fn:{encodedLabel}\">{encodedLabel}</a></sup>";
    }
    string IContextualHtmlMarkdownInline.RenderHtml(HtmlOptions options) {
        if (options?.GitHubFootnoteHtml != true) {
            return RenderHtml();
        }

        var state = HtmlRenderContext.Footnotes;
        if (state == null || !state.IsDefined(Label)) {
            return HtmlTextEncoder.Encode(RenderMarkdown(), options);
        }

        var reference = state.RegisterReference(Label);
        return "<sup class=\"footnote-ref\"><a href=\"#fn-"
               + HtmlTextEncoder.Encode(reference.EscapedLabel, options)
               + "\" id=\""
               + HtmlTextEncoder.Encode(reference.ReferenceId, options)
               + "\" data-footnote-ref>"
               + reference.Number.ToString()
               + "</a></sup>";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Label);
}
