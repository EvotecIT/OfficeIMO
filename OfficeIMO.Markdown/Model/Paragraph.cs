namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for a paragraph with plain text.
/// </summary>
public sealed class Paragraph : IMarkdownBlock {
    private readonly ParagraphBlock _p;
    /// <summary>Creates a paragraph from plain text.</summary>
    public Paragraph(string text) { _p = new ParagraphBlock(new InlineSequence().Text(text)); }
    /// <inheritdoc />
    public string RenderMarkdown() => _p.RenderMarkdown();
    /// <inheritdoc />
    public string RenderHtml() => _p.RenderHtml();
}
