namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for paragraphs composed of inline nodes.
/// </summary>
public sealed class ParagraphBuilder {
    internal InlineSequence Inlines { get; } = new InlineSequence();
    /// <summary>Appends plain text.</summary>
    public ParagraphBuilder Text(string text) { Inlines.Text(text); return this; }
    /// <summary>Appends a hyperlink.</summary>
    public ParagraphBuilder Link(string text, string url, string? title = null) { Inlines.Link(text, url, title); return this; }
    /// <summary>Appends bold text.</summary>
    public ParagraphBuilder Bold(string text) { Inlines.Bold(text); return this; }
    /// <summary>Appends italic text.</summary>
    public ParagraphBuilder Italic(string text) { Inlines.Italic(text); return this; }
    /// <summary>Appends strikethrough text.</summary>
    public ParagraphBuilder Strike(string text) { Inlines.Strike(text); return this; }
    /// <summary>Appends inline code.</summary>
    public ParagraphBuilder Code(string text) { Inlines.Code(text); return this; }
    /// <summary>Appends a linked image (e.g., a badge).</summary>
    public ParagraphBuilder ImageLink(string alt, string imageUrl, string linkUrl, string? title = null) { Inlines.ImageLink(alt, imageUrl, linkUrl, title); return this; }
}
