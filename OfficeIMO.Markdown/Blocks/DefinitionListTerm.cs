namespace OfficeIMO.Markdown;

/// <summary>
/// Semantic definition-list term with structured inline content.
/// </summary>
public sealed class DefinitionListTerm : MarkdownObject {
    private InlineSequence _inlines;
    internal string GenericAttributeConsumedWhitespace { get; set; } = string.Empty;

    /// <summary>Creates a definition-list term.</summary>
    public DefinitionListTerm(InlineSequence? inlines = null) {
        _inlines = inlines ?? new InlineSequence();
    }

    /// <summary>Inline content for the term.</summary>
    public InlineSequence Inlines {
        get => _inlines;
        set => _inlines = value ?? new InlineSequence();
    }

    /// <summary>Markdown representation of the term.</summary>
    public string Markdown {
        get {
            var markdown = Inlines.RenderMarkdown();
            if (Attributes.IsEmpty) {
                return markdown;
            }

            var separator = string.IsNullOrEmpty(GenericAttributeConsumedWhitespace)
                ? " "
                : GenericAttributeConsumedWhitespace;
            return markdown + separator + MarkdownAttributeBlockRenderer.RenderInlineTrailing(Attributes);
        }
    }

    /// <summary>Plain-text representation of the term.</summary>
    public string Text => InlinePlainText.Extract(Inlines);
}
