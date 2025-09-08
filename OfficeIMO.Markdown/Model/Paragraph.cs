namespace OfficeIMO.Markdown;

public sealed class Paragraph : IMarkdownBlock {
    private readonly ParagraphBlock _p;
    public Paragraph(string text) { _p = new ParagraphBlock(new InlineSequence().Text(text)); }
    public string RenderMarkdown() => _p.RenderMarkdown();
    public string RenderHtml() => _p.RenderHtml();
}

