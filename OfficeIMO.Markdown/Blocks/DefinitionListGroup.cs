namespace OfficeIMO.Markdown;

/// <summary>
/// Semantic definition-list group containing one or more shared terms and definition bodies.
/// </summary>
public sealed class DefinitionListGroup {
    private readonly List<InlineSequence> _terms = new List<InlineSequence>();
    private readonly List<DefinitionListDefinition> _definitions = new List<DefinitionListDefinition>();

    /// <summary>Terms that share the definitions in this group.</summary>
    public IReadOnlyList<InlineSequence> Terms => _terms;

    /// <summary>Definition bodies shared by the terms in this group.</summary>
    public IReadOnlyList<DefinitionListDefinition> Definitions => _definitions;

    /// <summary>Creates a grouped definition-list item.</summary>
    public DefinitionListGroup(
        IEnumerable<InlineSequence>? terms = null,
        IEnumerable<DefinitionListDefinition>? definitions = null) {
        if (terms != null) {
            foreach (var term in terms) {
                if (term != null) {
                    _terms.Add(term);
                }
            }
        }

        if (definitions != null) {
            foreach (var definition in definitions) {
                if (definition != null) {
                    _definitions.Add(definition);
                }
            }
        }
    }

    internal void AddTerm(InlineSequence? term) {
        if (term != null) {
            _terms.Add(term);
        }
    }

    internal void AddDefinition(DefinitionListDefinition? definition) {
        if (definition != null) {
            _definitions.Add(definition);
        }
    }
}
