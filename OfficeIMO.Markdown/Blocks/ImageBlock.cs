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
    /// <inheritdoc />
    public string? Caption { get; set; }

    /// <summary>Create an image block.</summary>
    public ImageBlock(string path, string? alt, string? title) {
        Path = path;
        Alt = alt;
        Title = title;
    }

    /// <inheritdoc />
    public string RenderMarkdown() {
        string alt = Alt ?? string.Empty;
        string title = string.IsNullOrEmpty(Title) ? string.Empty : " \"" + Title + "\"";
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.AppendLine($"![{alt}]({Path}{title})");
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    public string RenderHtml() {
        string alt = System.Net.WebUtility.HtmlEncode(Alt ?? string.Empty);
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        string src = System.Net.WebUtility.HtmlEncode(Path);
        string img = $"<img src=\"{src}\" alt=\"{alt}\"{title} />";
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return img + caption;
    }
}
