namespace OfficeIMO.Markdown;

public sealed class ParagraphBuilder {
    internal InlineSequence Inlines { get; } = new InlineSequence();
    public ParagraphBuilder Text(string text) { Inlines.Text(text); return this; }
    public ParagraphBuilder Link(string text, string url, string? title = null) { Inlines.Link(text, url, title); return this; }
}

