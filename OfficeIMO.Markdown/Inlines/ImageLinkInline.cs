namespace OfficeIMO.Markdown;

/// <summary>
/// Inline that renders a linked image, e.g. [![alt](img)](href).
/// Useful for badges (Shields.io).
/// </summary>
public sealed class ImageLinkInline {
    public string Alt { get; }
    public string ImageUrl { get; }
    public string LinkUrl { get; }
    public string? Title { get; }
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

