namespace OfficeIMO.Markdown;

/// <summary>
/// Typed definition list entry containing a term and one or more definition blocks.
/// </summary>
public sealed class DefinitionListEntry : MarkdownObject {
    private readonly DefinitionListDefinition _definition;
    private InlineSequence _term;
    private DefinitionListBlock? _owner;
    private DefinitionListGroup? _ownerGroup;
    private int _ownerTermIndex = -1;

    /// <summary>Inline content for the definition term.</summary>
    public InlineSequence Term {
        get => _term;
        set {
            var safeTerm = value ?? new InlineSequence();
            if (_owner != null) {
                _owner.ReplaceEntryTerm(this, safeTerm);
            } else {
                _term = safeTerm;
            }
        }
    }

    /// <summary>Structured definition content for this term.</summary>
    public List<IMarkdownBlock> DefinitionBlocks => _definition.Blocks;

    internal DefinitionListDefinition Definition => _definition;

    /// <summary>Creates a typed definition list entry.</summary>
    public DefinitionListEntry(InlineSequence? term = null, IEnumerable<IMarkdownBlock>? definitionBlocks = null)
        : this(term, new DefinitionListDefinition(definitionBlocks)) {
    }

    internal DefinitionListEntry(InlineSequence? term, DefinitionListDefinition? definition) {
        _term = term ?? new InlineSequence();
        _definition = definition ?? new DefinitionListDefinition();
    }

    internal void BindToDefinitionList(DefinitionListBlock owner, DefinitionListGroup group, int termIndex) {
        _owner = owner;
        _ownerGroup = group;
        _ownerTermIndex = termIndex;
    }

    internal bool IsBoundTo(DefinitionListGroup group, int termIndex) =>
        ReferenceEquals(_ownerGroup, group) && _ownerTermIndex == termIndex;

    internal void SetTermFromOwner(InlineSequence? term) {
        _term = term ?? new InlineSequence();
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
