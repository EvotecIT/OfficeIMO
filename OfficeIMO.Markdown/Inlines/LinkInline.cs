namespace OfficeIMO.Markdown;

/// <summary>
/// Hyperlink inline.
/// </summary>
public sealed class LinkInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline {
    /// <summary>Link text.</summary>
    public string Text { get; }
    /// <summary>Destination URL.</summary>
    public string Url { get; }
    /// <summary>Optional title shown as a tooltip in HTML.</summary>
    public string? Title { get; }
    /// <summary>Optional HTML target attribute preserved from richer sources.</summary>
    public string? LinkTarget { get; }
    /// <summary>Optional HTML rel attribute preserved from richer sources.</summary>
    public string? LinkRel { get; }

    // Optional richer label representation (produced by the reader). When present, RenderHtml/RenderMarkdown
    // uses it instead of the plain Text property.
    /// <summary>
    /// Optional richer inline representation of the link label when the reader parsed nested formatting.
    /// When present, renderers and converters can preserve label formatting instead of flattening to plain text.
    /// </summary>
    public InlineSequence? LabelInlines { get; }
    /// <summary>Creates a hyperlink inline.</summary>
    public LinkInline(string text, string url, string? title, string? linkTarget = null, string? linkRel = null) {
        Text = text ?? string.Empty;
        Url = url ?? string.Empty;
        Title = title;
        LinkTarget = linkTarget;
        LinkRel = linkRel;
    }

    internal LinkInline(InlineSequence label, string url, string? title, string? linkTarget = null, string? linkRel = null)
        : this(InlinePlainText.Extract(label), url, title, linkTarget, linkRel) {
        LabelInlines = label;
    }
    internal string RenderMarkdown() {
        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        if (LabelInlines != null) {
            return $"[{LabelInlines.RenderMarkdown()}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
        }
        return $"[{MarkdownEscaper.EscapeLinkText(Text)}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
    }
    internal string RenderHtml() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        var o = HtmlRenderContext.Options;
        if (!UrlOriginPolicy.IsAllowedHttpLink(o, Url)) {
            if (LabelInlines != null) return LabelInlines.RenderHtml();
            return System.Net.WebUtility.HtmlEncode(Text);
        }
        string extra = LinkHtmlAttributes.BuildLinkAttributes(o, Url, LinkTarget, LinkRel);
        if (LabelInlines != null) {
            return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{LabelInlines.RenderHtml()}</a>";
        }
        return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{System.Net.WebUtility.HtmlEncode(Text)}</a>";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => LabelInlines;
}
