namespace OfficeIMO.Markdown;

/// <summary>
/// Inline that renders a linked image, e.g. [![alt](img)](href).
/// Useful for badges (Shields.io).
/// </summary>
public sealed class ImageLinkInline {
    /// <summary>Alternative text for the image.</summary>
    public string Alt { get; }
    /// <summary>Image source URL (e.g., a Shields.io badge).</summary>
    public string ImageUrl { get; }
    /// <summary>Link target URL.</summary>
    public string LinkUrl { get; }
    /// <summary>Optional title.</summary>
    public string? Title { get; }
    /// <summary>Creates a linked image inline.</summary>
    public ImageLinkInline(string alt, string imageUrl, string linkUrl, string? title = null) {
        Alt = alt ?? string.Empty; ImageUrl = imageUrl ?? string.Empty; LinkUrl = linkUrl ?? string.Empty; Title = title;
    }
    internal string RenderMarkdown() {
        var title = string.IsNullOrEmpty(Title) ? string.Empty : " \"" + Title + "\"";
        return $"[![{Alt}]({ImageUrl}{title})]({LinkUrl})";
    }
    internal string RenderHtml() {
        var title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        return $"<a href=\"{System.Net.WebUtility.HtmlEncode(LinkUrl)}\"><img src=\"{System.Net.WebUtility.HtmlEncode(ImageUrl)}\" alt=\"{System.Net.WebUtility.HtmlEncode(Alt)}\"{title} /></a>";
    }
}
