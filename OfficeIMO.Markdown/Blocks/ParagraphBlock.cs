namespace OfficeIMO.Markdown;

public sealed class ParagraphBlock : IMarkdownBlock {
    public InlineSequence Inlines { get; }
    public ParagraphBlock(InlineSequence inlines) { Inlines = inlines; }
    public string RenderMarkdown() => Inlines.RenderMarkdown();
    public string RenderHtml() => $"<p>{Inlines.RenderHtml()}</p>";
}

