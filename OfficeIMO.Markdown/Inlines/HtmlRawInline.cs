namespace OfficeIMO.Markdown;

/// <summary>
/// Raw inline HTML passthrough. Rendering obeys the active <see cref="RawHtmlHandling"/> policy.
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
        var handling = options?.RawHtmlHandling ?? RawHtmlHandling.Allow;
        return handling switch {
            RawHtmlHandling.Allow => Html,
            RawHtmlHandling.Escape => System.Net.WebUtility.HtmlEncode(Html),
            RawHtmlHandling.Sanitize => RawHtmlSanitizer.Sanitize(Html),
            _ => string.Empty
        };
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) { }
}
