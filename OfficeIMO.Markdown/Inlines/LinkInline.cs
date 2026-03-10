namespace OfficeIMO.Markdown;

/// <summary>
/// Hyperlink inline.
/// </summary>
public sealed class LinkInline {
    /// <summary>Link text.</summary>
    public string Text { get; }
    /// <summary>Destination URL.</summary>
    public string Url { get; }
    /// <summary>Optional title shown as a tooltip in HTML.</summary>
    public string? Title { get; }

    // Optional richer label representation (produced by the reader). When present, RenderHtml/RenderMarkdown
    // uses it instead of the plain Text property.
    /// <summary>
    /// Optional richer inline representation of the link label when the reader parsed nested formatting.
    /// When present, renderers and converters can preserve label formatting instead of flattening to plain text.
    /// </summary>
    public InlineSequence? LabelInlines { get; }
    /// <summary>Creates a hyperlink inline.</summary>
    public LinkInline(string text, string url, string? title) { Text = text ?? string.Empty; Url = url ?? string.Empty; Title = title; }

    internal LinkInline(InlineSequence label, string url, string? title)
        : this(InlinePlainText.Extract(label), url, title) {
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
        string extra = LinkHtmlAttributes.BuildExternalLinkAttributes(o, Url);
        if (LabelInlines != null) {
            return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{LabelInlines.RenderHtml()}</a>";
        }
        return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{System.Net.WebUtility.HtmlEncode(Text)}</a>";
    }
}
