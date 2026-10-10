using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A native text story. Linked text frames can reference the same story; page layout is reported separately.</summary>
public sealed class PublisherTextStory {
    internal PublisherTextStory(uint id, string text, IReadOnlyList<OfficeRichTextParagraph> paragraphs) {
        Id = id; Text = text; Paragraphs = Array.AsReadOnly(paragraphs.ToArray());
    }
    /// <summary>Native story identifier used by publication text frames.</summary>
    public uint Id { get; }
    /// <summary>Complete recovered story text, with native carriage returns normalized to newlines.</summary>
    public string Text { get; }
    /// <summary>Paragraphs and styled runs recovered from the native text stream.</summary>
    public IReadOnlyList<OfficeRichTextParagraph> Paragraphs { get; }
}
