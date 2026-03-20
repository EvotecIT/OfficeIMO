namespace OfficeIMO.Markdown;

/// <summary>
/// Typed definition list entry containing a term and one or more definition blocks.
/// </summary>
public sealed class DefinitionListEntry : MarkdownObject {
    private readonly DefinitionListDefinition _definition;

    /// <summary>Inline content for the definition term.</summary>
    public InlineSequence Term { get; set; }

    /// <summary>Structured definition content for this term.</summary>
    public List<IMarkdownBlock> DefinitionBlocks => _definition.Blocks;

    internal DefinitionListDefinition Definition => _definition;

    /// <summary>Creates a typed definition list entry.</summary>
    public DefinitionListEntry(InlineSequence? term = null, IEnumerable<IMarkdownBlock>? definitionBlocks = null)
        : this(term, new DefinitionListDefinition(definitionBlocks)) {
    }

    internal DefinitionListEntry(InlineSequence? term, DefinitionListDefinition? definition) {
        Term = term ?? new InlineSequence();
        _definition = definition ?? new DefinitionListDefinition();
    }

    /// <summary>Markdown representation of the term.</summary>
    public string TermMarkdown => Term.RenderMarkdown();

    /// <summary>Markdown representation of the full definition body.</summary>
    public string DefinitionMarkdown => RenderDefinitionMarkdown();

    internal string RenderDefinitionMarkdown() {
        return _definition.RenderMarkdown();
    }

    internal string RenderDefinitionHtml() {
        return _definition.RenderHtml();
    }
}
