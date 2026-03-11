namespace OfficeIMO.Markdown;

/// <summary>
/// Parsed inline representation of a definition list entry.
/// </summary>
public readonly struct DefinitionListInlineItem {
    /// <summary>Inline content for the definition term.</summary>
    public InlineSequence Term { get; }

    /// <summary>Inline content for the definition value.</summary>
    public InlineSequence Definition { get; }

    /// <summary>Creates a parsed definition list entry.</summary>
    public DefinitionListInlineItem(InlineSequence term, InlineSequence definition) {
        Term = term ?? new InlineSequence();
        Definition = definition ?? new InlineSequence();
    }
}
