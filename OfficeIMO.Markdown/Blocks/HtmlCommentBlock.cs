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

    string IMarkdownBlock.RenderHtml() => Comment;
}
