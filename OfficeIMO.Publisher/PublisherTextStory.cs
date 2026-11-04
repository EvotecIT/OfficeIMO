using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A complete current text story. Linked text frames can reference the same story; page layout is reported separately.</summary>
public sealed class PublisherTextStory {
    internal PublisherTextStory(uint id, string text, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        bool? hasFinalParagraphTerminator = null, bool? automaticHyphenation = null) {
        Id = id; Text = text;
        Paragraphs = Array.AsReadOnly(paragraphs.Select(paragraph => {
            if (paragraph.UsesPublisherFontMetrics) return paragraph;
            OfficeRichTextParagraph copy = paragraph.WithTabStops(paragraph.TabStops);
            copy.UsesPublisherFontMetrics = true;
            return copy;
        }).ToArray());
        HasFinalParagraphTerminator = hasFinalParagraphTerminator ?? text.EndsWith("\n", StringComparison.Ordinal);
        AutomaticHyphenation = automaticHyphenation;
    }
    // This source convention survives clearing the content. Inferring it from
    // the current empty text would change subsequent story/frame character ranges.
    internal bool HasFinalParagraphTerminator { get; }
    /// <summary>Stable publication story identifier used by text frames or tables.</summary>
    public uint Id { get; }
    /// <summary>Complete current story text, with carriage returns normalized to newlines and paragraphs separated by newlines. Authored ordinary stories have no implicit final terminator; imported stories retain their source convention. List labels remain separate from this text.</summary>
    public string Text { get; }
    /// <summary>Current paragraphs and styled runs, initially recovered from the native text stream.</summary>
    public IReadOnlyList<OfficeRichTextParagraph> Paragraphs { get; }
    /// <summary>Native story automatic-hyphenation setting, including the qualified enabled default when its catalog entry omits the flag. Null means the setting is unavailable or unassessed. Managed display uses supported shared patterns while preserving the original source text.</summary>
    public bool? AutomaticHyphenation { get; }
}
