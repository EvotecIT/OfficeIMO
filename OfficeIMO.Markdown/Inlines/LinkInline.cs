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
        string? autolinkMarkdown = TryRenderAutolinkMarkdown();
        if (autolinkMarkdown != null) {
            return autolinkMarkdown;
        }

        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        if (LabelInlines != null) {
            return $"[{LabelInlines.RenderMarkdown()}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
        }
        return $"[{MarkdownEscaper.EscapeLinkText(Text)}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
    }

    private string? TryRenderAutolinkMarkdown() {
        if (!string.IsNullOrEmpty(Title) || LabelInlines != null) {
            return null;
        }

        if (!MarkdownInlineMetadataSourceSpans.GetLinkTargetSpan(this).HasValue) {
            return null;
        }

        string? openingMarker = MarkdownInlineMetadataSourceSpans.GetOpeningMarker(this);
        string? closingMarker = MarkdownInlineMetadataSourceSpans.GetClosingMarker(this);
        string? separatorMarker = MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(this);
        if (string.Equals(openingMarker, "<", StringComparison.Ordinal) &&
            string.Equals(closingMarker, ">", StringComparison.Ordinal) &&
            string.IsNullOrEmpty(separatorMarker)) {
            return "<" + Text + ">";
        }

        if (!string.IsNullOrEmpty(openingMarker) ||
            !string.IsNullOrEmpty(closingMarker) ||
            !string.IsNullOrEmpty(separatorMarker)) {
            return null;
        }

        string? autolinkLiteral = MarkdownInlineMetadataSourceSpans.GetAutolinkLiteral(this);
        if (!string.IsNullOrEmpty(autolinkLiteral)) {
            return autolinkLiteral;
        }

        if (string.Equals(Text, Url, StringComparison.Ordinal) ||
            (Url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) &&
             string.Equals(Text, Url.Substring("mailto:".Length), StringComparison.Ordinal)) ||
            (Url.StartsWith("tel:", StringComparison.OrdinalIgnoreCase) &&
             string.Equals(Text, Url.Substring("tel:".Length), StringComparison.Ordinal)) ||
            (Url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
             string.Equals(Text, Url.Substring("http://".Length), StringComparison.Ordinal)) ||
            (Url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
             string.Equals(Text, Url.Substring("https://".Length), StringComparison.Ordinal))) {
            return Url.StartsWith("tel:", StringComparison.OrdinalIgnoreCase) ? Url : Text;
        }

        return null;
    }
    internal string RenderHtml() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        var o = HtmlRenderContext.Options;
        if (!UrlOriginPolicy.IsAllowedHttpLink(o, Url)) {
            if (LabelInlines != null) return LabelInlines.RenderHtml();
            return HtmlTextEncoder.Encode(Text);
        }
        string extra = LinkHtmlAttributes.BuildLinkAttributes(o, Url, LinkTarget, LinkRel);
        if (LabelInlines != null) {
            return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{LabelInlines.RenderHtml()}</a>";
        }
        return $"<a href=\"{HtmlAttributeUrlEncoder.Encode(Url)}\"{title}{extra}>{HtmlTextEncoder.Encode(Text)}</a>";
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => LabelInlines;
}
