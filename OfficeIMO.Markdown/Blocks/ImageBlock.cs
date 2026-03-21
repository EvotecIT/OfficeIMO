namespace OfficeIMO.Markdown;

/// <summary>
/// Image block with optional title and caption.
/// </summary>
public sealed class ImageBlock : MarkdownBlock, IMarkdownBlock, ICaptionable, ISyntaxMarkdownBlock {
    /// <summary>Image source path or URL.</summary>
    public string Path { get; }
    /// <summary>Optional hyperlink target wrapping the image.</summary>
    public string? LinkUrl { get; set; }
    /// <summary>Optional hyperlink title wrapping the image.</summary>
    public string? LinkTitle { get; set; }
    /// <summary>Optional hyperlink target attribute wrapping the image in HTML.</summary>
    public string? LinkTarget { get; set; }
    /// <summary>Optional hyperlink rel attribute wrapping the image in HTML.</summary>
    public string? LinkRel { get; set; }
    /// <summary>Alternative text.</summary>
    public string? Alt { get; }
    /// <summary>Optional title attribute.</summary>
    public string? Title { get; }
    /// <summary>Optional width hint (points/pixels as provided).</summary>
    public double? Width { get; set; }
    /// <summary>Optional height hint.</summary>
    public double? Height { get; set; }
    /// <summary>Optional fallback image URL used when rendering preserved picture sources back to HTML.</summary>
    public string? PictureFallbackPath { get; set; }
    /// <summary>HTML-only responsive picture source metadata preserved from imported &lt;picture&gt; elements.</summary>
    public IList<ImagePictureSource> PictureSources { get; } = new List<ImagePictureSource>();
    /// <inheritdoc />
    public string? Caption { get; set; }
    internal MarkdownSourceSpan? AltSyntaxSpan { get; private set; }
    internal MarkdownSourceSpan? SourceSyntaxSpan { get; private set; }
    internal MarkdownSourceSpan? TitleSyntaxSpan { get; private set; }
    internal MarkdownSourceSpan? LinkTargetSyntaxSpan { get; private set; }
    internal MarkdownSourceSpan? LinkTitleSyntaxSpan { get; private set; }

    /// <summary>Create an image block.</summary>
    public ImageBlock(string path, string? alt, string? title)
        : this(path, alt, title, null, null, null, null, null, null) {
    }

    /// <summary>Create an image block with optional size hints.</summary>
    public ImageBlock(string path, string? alt = null, string? title = null, double? width = null, double? height = null, string? linkUrl = null, string? linkTitle = null, string? linkTarget = null, string? linkRel = null) {
        Path = path;
        Alt = alt;
        Title = title;
        Width = width;
        Height = height;
        LinkUrl = linkUrl;
        LinkTitle = linkTitle;
        LinkTarget = linkTarget;
        LinkRel = linkRel;
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        var imageRenderingMode = MarkdownRenderContext.Options?.ImageRenderingMode ?? MarkdownImageRenderingMode.RichMarkdown;
        if (imageRenderingMode == MarkdownImageRenderingMode.Html) {
            return ((IMarkdownBlock)this).RenderHtml();
        }

        string alt = MarkdownEscaper.EscapeImageAlt(Alt ?? string.Empty);
        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        string imageMarkdown = $"![{alt}]({MarkdownEscaper.EscapeImageSrc(Path)}{title})";
        if (!string.IsNullOrWhiteSpace(LinkUrl)) {
            sb.Append($"[{imageMarkdown}]({MarkdownEscaper.EscapeLinkUrl(LinkUrl!)}{MarkdownEscaper.FormatOptionalTitle(LinkTitle)})");
        } else {
            sb.Append(imageMarkdown);
        }
        if (imageRenderingMode == MarkdownImageRenderingMode.RichMarkdown && (Width != null || Height != null)) {
            var w = Width != null ? $"width={Width.Value}" : string.Empty;
            var h = Height != null ? $"height={Height.Value}" : string.Empty;
            var sep = (w != string.Empty && h != string.Empty) ? " " : string.Empty;
            sb.Append($"{{{w}{sep}{h}}}");
        }
        sb.AppendLine();
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        string alt = System.Net.WebUtility.HtmlEncode(Alt ?? string.Empty);
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        string size = string.Empty;
        if (Width != null) size += $" width=\"{Width.Value}\"";
        if (Height != null) size += $" height=\"{Height.Value}\"";
        var o = HtmlRenderContext.Options;
        if (!UrlOriginPolicy.IsAllowedHttpImage(o, Path)) {
            string captionBlocked = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
            return ImageHtmlAttributes.BuildBlockedPlaceholder(Alt) + captionBlocked;
        }
        var extra = ImageHtmlAttributes.BuildImageAttributes(o, Path);
        string img = $"<img src=\"{HtmlAttributeUrlEncoder.Encode(GetRenderedFallbackImagePath(o))}\" alt=\"{alt}\"{title}{size}{extra} />";
        if (PictureSources.Count > 0) {
            img = BuildPictureHtml(o, img);
        }
        if (!string.IsNullOrWhiteSpace(LinkUrl) && UrlOriginPolicy.IsAllowedHttpLink(o, LinkUrl!)) {
            string linkExtra = BuildLinkHtmlAttributes(o, LinkUrl!, LinkTarget, LinkRel);
            string linkTitle = string.IsNullOrEmpty(LinkTitle) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(LinkTitle!)}\"";
            img = $"<a href=\"{HtmlAttributeUrlEncoder.Encode(LinkUrl!)}\"{linkTitle}{linkExtra}>{img}</a>";
        }
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return img + caption;
    }

    private string GetRenderedFallbackImagePath(HtmlOptions? options) {
        if (!string.IsNullOrWhiteSpace(PictureFallbackPath) && UrlOriginPolicy.IsAllowedHttpImage(options, PictureFallbackPath!)) {
            return PictureFallbackPath!;
        }

        return Path;
    }

    private string BuildPictureHtml(HtmlOptions? options, string imgHtml) {
        var sb = new System.Text.StringBuilder();
        sb.Append("<picture>");
        for (int i = 0; i < PictureSources.Count; i++) {
            var source = PictureSources[i];
            if (source == null || string.IsNullOrWhiteSpace(source.Path) || !UrlOriginPolicy.IsAllowedHttpImage(options, source.Path)) {
                continue;
            }

            string srcSet = NormalizeAttributeValue(source.SrcSet) ?? source.Path;
            sb.Append("<source srcset=\"")
                .Append(HtmlAttributeUrlEncoder.EncodeSrcSet(srcSet))
                .Append('"');
            AppendAttribute(sb, "media", NormalizeAttributeValue(source.Media));
            AppendAttribute(sb, "type", NormalizeAttributeValue(source.Type));
            AppendAttribute(sb, "sizes", NormalizeAttributeValue(source.Sizes));
            sb.Append(" />");
        }

        sb.Append(imgHtml);
        sb.Append("</picture>");
        return sb.ToString();
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode>();
        if (!string.IsNullOrEmpty(Alt)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageAlt, AltSyntaxSpan, Alt));
        }

        nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageSource, SourceSyntaxSpan, Path));

        if (!string.IsNullOrEmpty(LinkUrl)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageLinkTarget, LinkTargetSyntaxSpan, LinkUrl));
        }

        if (!string.IsNullOrEmpty(LinkTitle)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageLinkTitle, LinkTitleSyntaxSpan, LinkTitle));
        }

        if (!string.IsNullOrEmpty(LinkTarget)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageLinkHtmlTarget, span, LinkTarget));
        }

        if (!string.IsNullOrEmpty(LinkRel)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageLinkHtmlRel, span, LinkRel));
        }

        if (!string.IsNullOrEmpty(Title)) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ImageTitle, TitleSyntaxSpan, Title));
        }

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Image,
            span,
            ((IMarkdownBlock)this).RenderMarkdown(),
            nodes,
            this);
    }

    internal void SetMarkdownSyntaxMetadataSpans(
        MarkdownSourceSpan? altSpan,
        MarkdownSourceSpan? sourceSpan,
        MarkdownSourceSpan? titleSpan,
        MarkdownSourceSpan? linkTargetSpan,
        MarkdownSourceSpan? linkTitleSpan) {
        AltSyntaxSpan = altSpan;
        SourceSyntaxSpan = sourceSpan;
        TitleSyntaxSpan = titleSpan;
        LinkTargetSyntaxSpan = linkTargetSpan;
        LinkTitleSyntaxSpan = linkTitleSpan;
    }

    internal void CopyMarkdownSyntaxMetadataSpansFrom(ImageBlock? source) {
        if (source == null) {
            return;
        }

        SetMarkdownSyntaxMetadataSpans(
            source.AltSyntaxSpan,
            source.SourceSyntaxSpan,
            source.TitleSyntaxSpan,
            source.LinkTargetSyntaxSpan,
            source.LinkTitleSyntaxSpan);
    }

    private static string BuildLinkHtmlAttributes(HtmlOptions? options, string url, string? explicitTarget, string? explicitRel) {
        var generated = LinkHtmlAttributes.BuildExternalLinkAttributes(options, url);
        var target = NormalizeAttributeValue(explicitTarget);
        var rel = NormalizeAttributeValue(explicitRel);
        var referrerPolicy = ExtractAttributeValue(generated, "referrerpolicy");

        if (string.IsNullOrEmpty(target)) {
            target = ExtractAttributeValue(generated, "target");
        }

        if (string.IsNullOrEmpty(rel)) {
            rel = ExtractAttributeValue(generated, "rel");
        }

        rel = HardenRelForTarget(target, rel);

        var sb = new System.Text.StringBuilder();
        AppendAttribute(sb, "target", target);
        AppendAttribute(sb, "rel", rel);
        AppendAttribute(sb, "referrerpolicy", referrerPolicy);
        return sb.ToString();
    }

    private static string? NormalizeAttributeValue(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string trimmed = value!.Trim();
        return trimmed.Length == 0 ? null : trimmed;
    }

    private static string? ExtractAttributeValue(string htmlAttributes, string attributeName) {
        if (string.IsNullOrWhiteSpace(htmlAttributes) || string.IsNullOrWhiteSpace(attributeName)) {
            return null;
        }

        string pattern = attributeName + "=\"";
        int start = htmlAttributes.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return null;
        }

        start += pattern.Length;
        int end = htmlAttributes.IndexOf('"', start);
        if (end < 0) {
            return null;
        }

        string encoded = htmlAttributes.Substring(start, end - start);
        return System.Net.WebUtility.HtmlDecode(encoded);
    }

    private static string? HardenRelForTarget(string? target, string? rel) {
        if (!string.Equals(target, "_blank", StringComparison.OrdinalIgnoreCase)) {
            return rel;
        }

        var normalizedRel = NormalizeAttributeValue(rel) ?? string.Empty;
        normalizedRel = AppendTokenIfMissing(normalizedRel, "noopener");
        normalizedRel = AppendTokenIfMissing(normalizedRel, "noreferrer");
        return normalizedRel;
    }

    private static string AppendTokenIfMissing(string value, string token) {
        if (string.IsNullOrWhiteSpace(token)) {
            return value;
        }

        var tokenParts = (value ?? string.Empty)
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokenParts.Any(existing => string.Equals(existing, token, StringComparison.OrdinalIgnoreCase))) {
            return value ?? string.Empty;
        }

        return string.IsNullOrWhiteSpace(value) ? token : value + " " + token;
    }

    private static void AppendAttribute(System.Text.StringBuilder sb, string attributeName, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        sb.Append(' ')
            .Append(attributeName)
            .Append("=\"")
            .Append(System.Net.WebUtility.HtmlEncode(value))
            .Append('"');
    }
}
