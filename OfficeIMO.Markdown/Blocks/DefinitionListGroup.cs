namespace OfficeIMO.Markdown;

/// <summary>
/// Semantic definition-list group containing one or more shared terms and definition bodies.
/// </summary>
public sealed class DefinitionListGroup : MarkdownObject {
    private readonly List<DefinitionListTerm> _terms = new List<DefinitionListTerm>();
    private readonly IReadOnlyList<InlineSequence> _termInlines;
    private readonly List<DefinitionListDefinition> _definitions = new List<DefinitionListDefinition>();

    /// <summary>Creates an empty definition-list group.</summary>
    public DefinitionListGroup() {
        _termInlines = new DefinitionListTermInlineList(_terms);
    }

    /// <summary>Structured terms that share the definitions in this group.</summary>
    public IReadOnlyList<DefinitionListTerm> TermItems => _terms;

    /// <summary>Compatibility inline view of the terms that share the definitions in this group.</summary>
    public IReadOnlyList<InlineSequence> Terms => _termInlines;

    /// <summary>Definition bodies shared by the terms in this group.</summary>
    public IReadOnlyList<DefinitionListDefinition> Definitions => _definitions;

    /// <summary>Creates a grouped definition-list item.</summary>
    public DefinitionListGroup(
        IEnumerable<InlineSequence>? terms = null,
        IEnumerable<DefinitionListDefinition>? definitions = null)
        : this() {
        if (terms != null) {
            foreach (var term in terms) {
                if (term != null) {
                    _terms.Add(new DefinitionListTerm(term));
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
            _terms.Add(new DefinitionListTerm(term));
        }
    }

    internal void AddTerm(DefinitionListTerm? term) {
        if (term != null) {
            _terms.Add(term);
        }
    }

    internal void ReplaceTerm(int index, InlineSequence? term) {
        if (index < 0 || index >= _terms.Count) {
            return;
        }

        _terms[index].Inlines = term ?? new InlineSequence();
    }

    internal void AddDefinition(DefinitionListDefinition? definition) {
        if (definition != null) {
            _definitions.Add(definition);
        }
    }

    private sealed class DefinitionListTermInlineList : IReadOnlyList<InlineSequence> {
        private readonly IReadOnlyList<DefinitionListTerm> _terms;

        internal DefinitionListTermInlineList(IReadOnlyList<DefinitionListTerm> terms) {
            _terms = terms;
        }

        public int Count => _terms.Count;

        public InlineSequence this[int index] => _terms[index].Inlines;

        public IEnumerator<InlineSequence> GetEnumerator() {
            for (int i = 0; i < _terms.Count; i++) {
                yield return _terms[i].Inlines;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
