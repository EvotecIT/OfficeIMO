namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how images are emitted when rendering a typed document back to markdown text.
/// </summary>
public enum MarkdownImageRenderingMode {
    /// <summary>
    /// Emit OfficeIMO-flavored markdown, including attribute-list style size hints on image blocks.
    /// </summary>
    RichMarkdown = 0,

    /// <summary>
    /// Emit broadly compatible markdown and omit OfficeIMO-only image size suffixes.
    /// </summary>
    PortableMarkdown = 1,

    /// <summary>
    /// Emit raw HTML for images so width, height, and richer HTML metadata round-trip exactly.
    /// </summary>
    Html = 2
}
