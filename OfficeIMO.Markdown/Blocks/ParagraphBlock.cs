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
    string IMarkdownBlock.RenderMarkdown() => Inlines.RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => $"<p>{Inlines.RenderHtml()}</p>";
}
