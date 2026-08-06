namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for paragraphs composed of inline nodes.
/// </summary>
public sealed class MarkdownParagraphBuilder {
    internal InlineSequence Inlines { get; } = new InlineSequence();
    /// <summary>Appends plain text.</summary>
    public MarkdownParagraphBuilder Text(string text) { Inlines.Text(text); return this; }
    /// <summary>Appends a hyperlink.</summary>
    public MarkdownParagraphBuilder Link(string text, string url, string? title = null) { Inlines.Link(text, url, title); return this; }
    /// <summary>Appends bold text.</summary>
    public MarkdownParagraphBuilder Bold(string text) { Inlines.Bold(text); return this; }
    /// <summary>Appends italic text.</summary>
    public MarkdownParagraphBuilder Italic(string text) { Inlines.Italic(text); return this; }
    /// <summary>Appends strikethrough text.</summary>
    public MarkdownParagraphBuilder Strike(string text) { Inlines.Strike(text); return this; }
    /// <summary>Appends inline code.</summary>
    public MarkdownParagraphBuilder Code(string text) { Inlines.Code(text); return this; }
    /// <summary>Appends underlined text.</summary>
    public MarkdownParagraphBuilder Underline(string text) { Inlines.Underline(text); return this; }
    /// <summary>Appends a linked image (e.g., a badge).</summary>
    public MarkdownParagraphBuilder ImageLink(string alt, string imageUrl, string linkUrl, string? title = null) { Inlines.ImageLink(alt, imageUrl, linkUrl, title); return this; }
}
