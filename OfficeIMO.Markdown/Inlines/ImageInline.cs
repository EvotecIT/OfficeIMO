namespace OfficeIMO.Markdown;

/// <summary>
/// Standalone inline image: ![alt](src "title").
/// </summary>
public sealed class ImageInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Alternate text for the image.</summary>
    public string Alt { get; }
    /// <summary>Plain-text alternate text used for HTML rendering and text extraction.</summary>
    public string PlainAlt { get; }
    /// <summary>Image source URL or data URI.</summary>
    public string Src { get; }
    /// <summary>Optional title attribute shown as tooltip in HTML.</summary>
    public string? Title { get; }
    /// <summary>Creates a new inline image.</summary>
    public ImageInline(string alt, string src, string? title = null, string? plainAlt = null) {
        Alt = alt ?? string.Empty;
        PlainAlt = plainAlt ?? Alt;
        Src = src ?? string.Empty;
        Title = title;
    }
    internal string RenderMarkdown() {
        if ((MarkdownRenderContext.Options?.ImageRenderingMode ?? MarkdownImageRenderingMode.RichMarkdown) == MarkdownImageRenderingMode.Html) {
            return RenderHtml();
        }

        var title = MarkdownEscaper.FormatOptionalTitle(Title);
        return $"![{MarkdownEscaper.EscapeImageAlt(Alt)}]({MarkdownEscaper.EscapeImageSrc(Src)}{title})";
    }
    internal string RenderHtml() {
        var titleAttr = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title)}\"";
        var o = HtmlRenderContext.Options;
        if (!UrlOriginPolicy.IsAllowedHttpImage(o, Src)) {
            return ImageHtmlAttributes.BuildBlockedPlaceholder(PlainAlt);
        }
        var extra = ImageHtmlAttributes.BuildImageAttributes(o, Src);
        return $"<img src=\"{HtmlAttributeUrlEncoder.Encode(Src)}\" alt=\"{System.Net.WebUtility.HtmlEncode(PlainAlt)}\"{titleAttr}{extra} />";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(PlainAlt);
}
