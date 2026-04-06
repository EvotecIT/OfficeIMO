namespace OfficeIMO.Markdown;

/// <summary>
/// Inline that renders a linked image, e.g. [![alt](img)](href).
/// Useful for badges (Shields.io).
/// </summary>
public sealed class ImageLinkInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Alternative text for the image.</summary>
    public string Alt { get; }
    /// <summary>Plain-text alternate text used for HTML rendering and text extraction.</summary>
    public string PlainAlt { get; }
    /// <summary>Image source URL (e.g., a Shields.io badge).</summary>
    public string ImageUrl { get; }
    /// <summary>Link target URL.</summary>
    public string LinkUrl { get; }
    /// <summary>Optional image title.</summary>
    public string? Title { get; }
    /// <summary>Optional hyperlink title.</summary>
    public string? LinkTitle { get; }
    /// <summary>Creates a linked image inline.</summary>
    public ImageLinkInline(string alt, string imageUrl, string linkUrl, string? title = null, string? linkTitle = null, string? plainAlt = null) {
        Alt = alt ?? string.Empty;
        PlainAlt = plainAlt ?? Alt;
        ImageUrl = imageUrl ?? string.Empty;
        LinkUrl = linkUrl ?? string.Empty;
        Title = title;
        LinkTitle = linkTitle;
    }
    internal string RenderMarkdown() {
        if ((MarkdownRenderContext.Options?.ImageRenderingMode ?? MarkdownImageRenderingMode.RichMarkdown) == MarkdownImageRenderingMode.Html) {
            return RenderHtml();
        }

        var title = MarkdownEscaper.FormatOptionalTitle(Title);
        var linkTitle = MarkdownEscaper.FormatOptionalTitle(LinkTitle);
        return $"[![{MarkdownEscaper.EscapeImageAlt(Alt)}]({MarkdownEscaper.EscapeImageSrc(ImageUrl)}{title})]({MarkdownEscaper.EscapeLinkUrl(LinkUrl)}{linkTitle})";
    }
    internal string RenderHtml() {
        var title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        var linkTitle = string.IsNullOrEmpty(LinkTitle) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(LinkTitle!)}\"";
        var o = HtmlRenderContext.Options;
        bool linkAllowed = UrlOriginPolicy.IsAllowedHttpLink(o, LinkUrl);
        bool imageAllowed = UrlOriginPolicy.IsAllowedHttpImage(o, ImageUrl);

        var imgExtra = imageAllowed ? ImageHtmlAttributes.BuildImageAttributes(o, ImageUrl) : string.Empty;
        var extra = linkAllowed ? LinkHtmlAttributes.BuildExternalLinkAttributes(o, LinkUrl) : string.Empty;

        if (!linkAllowed && !imageAllowed) return ImageHtmlAttributes.BuildBlockedPlaceholder(PlainAlt);
        if (!imageAllowed && linkAllowed) {
            return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(LinkUrl)}\"{linkTitle}{extra}>{System.Net.WebUtility.HtmlEncode(PlainAlt)}</a>";
        }
        if (imageAllowed && !linkAllowed) {
            return $"<img src=\"{HtmlAttributeUrlEncoder.Encode(ImageUrl)}\" alt=\"{System.Net.WebUtility.HtmlEncode(PlainAlt)}\"{title}{imgExtra} />";
        }

        return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(LinkUrl)}\"{linkTitle}{extra}><img src=\"{HtmlAttributeUrlEncoder.Encode(ImageUrl)}\" alt=\"{System.Net.WebUtility.HtmlEncode(PlainAlt)}\"{title}{imgExtra} /></a>";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(PlainAlt);
}
