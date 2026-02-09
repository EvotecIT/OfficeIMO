namespace OfficeIMO.Markdown;

/// <summary>
/// Image block with optional title and caption.
/// </summary>
public sealed class ImageBlock : IMarkdownBlock, ICaptionable {
    /// <summary>Image source path or URL.</summary>
    public string Path { get; }
    /// <summary>Alternative text.</summary>
    public string? Alt { get; }
    /// <summary>Optional title attribute.</summary>
    public string? Title { get; }
    /// <summary>Optional width hint (points/pixels as provided).</summary>
    public double? Width { get; set; }
    /// <summary>Optional height hint.</summary>
    public double? Height { get; set; }
    /// <inheritdoc />
    public string? Caption { get; set; }

    /// <summary>Create an image block.</summary>
    public ImageBlock(string path, string? alt, string? title)
        : this(path, alt, title, null, null) {
    }

    /// <summary>Create an image block with optional size hints.</summary>
    public ImageBlock(string path, string? alt = null, string? title = null, double? width = null, double? height = null) {
        Path = path;
        Alt = alt;
        Title = title;
        Width = width;
        Height = height;
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        string alt = MarkdownEscaper.EscapeImageAlt(Alt ?? string.Empty);
        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.Append($"![{alt}]({MarkdownEscaper.EscapeImageSrc(Path)}{title})");
        if (Width != null || Height != null) {
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
        string src = System.Net.WebUtility.HtmlEncode(Path);
        string size = string.Empty;
        if (Width != null) size += $" width=\"{Width.Value}\"";
        if (Height != null) size += $" height=\"{Height.Value}\"";
        var o = HtmlRenderContext.Options;
        if (!UrlOriginPolicy.IsAllowedHttpImage(o, Path)) {
            string captionBlocked = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
            return ImageHtmlAttributes.BuildBlockedPlaceholder(Alt) + captionBlocked;
        }
        var extra = ImageHtmlAttributes.BuildImageAttributes(o, Path);
        string img = $"<img src=\"{src}\" alt=\"{alt}\"{title}{size}{extra} />";
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return img + caption;
    }
}
