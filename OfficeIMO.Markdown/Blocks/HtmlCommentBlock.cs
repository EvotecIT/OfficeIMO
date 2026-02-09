namespace OfficeIMO.Markdown;

/// <summary>
/// Represents an HTML comment preserved as a top-level Markdown block.
/// </summary>
public sealed class HtmlCommentBlock : IMarkdownBlock {
    /// <summary>Gets the raw HTML comment text, including the comment delimiters.</summary>
    public string Comment { get; }

    /// <summary>Initializes a new instance of the <see cref="HtmlCommentBlock"/> class.</summary>
    /// <param name="comment">HTML comment content to preserve verbatim.</param>
    public HtmlCommentBlock(string comment) {
        Comment = comment ?? string.Empty;
    }

    string IMarkdownBlock.RenderMarkdown() => Comment;

    string IMarkdownBlock.RenderHtml() {
        var o = HtmlRenderContext.Options;
        var handling = o?.RawHtmlHandling ?? RawHtmlHandling.Allow;
        return handling switch {
            RawHtmlHandling.Allow => Comment,
            RawHtmlHandling.Escape => "<pre class=\"md-raw-html\"><code>" + System.Net.WebUtility.HtmlEncode(Comment) + "</code></pre>",
            RawHtmlHandling.Sanitize => RawHtmlSanitizer.Sanitize(Comment),
            _ => string.Empty
        };
    }
}
