namespace OfficeIMO.Markdown;

/// <summary>
/// Standalone inline image: ![alt](src "title").
/// </summary>
public sealed class ImageInline {
    /// <summary>Alternate text for the image.</summary>
    public string Alt { get; }
    /// <summary>Image source URL or data URI.</summary>
    public string Src { get; }
    /// <summary>Optional title attribute shown as tooltip in HTML.</summary>
    public string? Title { get; }
    /// <summary>Creates a new inline image.</summary>
    public ImageInline(string alt, string src, string? title = null) { Alt = alt; Src = src; Title = title; }
    internal string RenderMarkdown() {
        var title = string.IsNullOrEmpty(Title) ? string.Empty : $" \"{Title}\"";
        return $"![{Alt}]({Src}{title})";
    }
    internal string RenderHtml() {
        var titleAttr = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title)}\"";
        return $"<img src=\"{System.Net.WebUtility.HtmlEncode(Src)}\" alt=\"{System.Net.WebUtility.HtmlEncode(Alt)}\"{titleAttr} />";
    }
}
