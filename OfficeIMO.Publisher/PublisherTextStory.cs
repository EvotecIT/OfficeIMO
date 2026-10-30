using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A native text story. Linked text frames can reference the same story; page layout is reported separately.</summary>
public sealed class PublisherTextStory {
    internal PublisherTextStory(uint id, string text, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        bool? hasFinalParagraphTerminator = null) {
        Id = id; Text = text; Paragraphs = Array.AsReadOnly(paragraphs.ToArray());
        HasFinalParagraphTerminator = hasFinalParagraphTerminator ?? text.EndsWith("\n", StringComparison.Ordinal);
    }
    // This source convention survives clearing the content. Inferring it from
    // the current empty text would change subsequent story/frame character ranges.
    internal bool HasFinalParagraphTerminator { get; }
    /// <summary>Native story identifier used by publication text frames.</summary>
    public uint Id { get; }
    /// <summary>Complete current story text, with native carriage returns normalized to newlines.</summary>
    public string Text { get; }
    /// <summary>Current paragraphs and styled runs, initially recovered from the native text stream.</summary>
    public IReadOnlyList<OfficeRichTextParagraph> Paragraphs { get; }
}
