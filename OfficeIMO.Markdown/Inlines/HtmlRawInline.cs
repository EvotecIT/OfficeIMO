namespace OfficeIMO.Markdown;

/// <summary>
/// Raw inline HTML passthrough. Rendering obeys the active <see cref="MarkdownRawHtmlHandling"/> policy.
/// </summary>
public sealed class HtmlRawInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Raw HTML fragment.</summary>
    public string Html { get; }

    /// <summary>Creates a new raw inline HTML node.</summary>
    public HtmlRawInline(string html) {
        Html = html ?? string.Empty;
    }

    internal string RenderMarkdown() => Html;

    internal string RenderHtml() {
        var options = HtmlRenderContext.Options;
        var handling = options?.RawHtmlHandling ?? MarkdownRawHtmlHandling.Allow;
        string rendered = handling switch {
            MarkdownRawHtmlHandling.Allow => options?.GitHubHtmlTagFilter == true ? GitHubHtmlTagFilter.Apply(Html) : Html,
            MarkdownRawHtmlHandling.Escape => HtmlTextEncoder.Encode(Html, options),
            MarkdownRawHtmlHandling.Sanitize => RawHtmlSanitizer.Sanitize(Html, options),
            _ => string.Empty
        };
        options?.RecordRawHtmlChange(Html, rendered, "RawHtmlInline");
        return rendered;
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) { }
}
