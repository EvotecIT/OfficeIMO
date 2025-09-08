namespace OfficeIMO.Markdown;

/// <summary>
/// Paragraph block containing a sequence of inline nodes.
/// </summary>
public sealed class ParagraphBlock : IMarkdownBlock {
    /// <summary>Inline content within this paragraph.</summary>
    public InlineSequence Inlines { get; }
    /// <summary>Creates a paragraph block.</summary>
    public ParagraphBlock(InlineSequence inlines) { Inlines = inlines; }
    /// <inheritdoc />
    public string RenderMarkdown() => Inlines.RenderMarkdown();
    /// <inheritdoc />
    public string RenderHtml() => $"<p>{Inlines.RenderHtml()}</p>";
}
