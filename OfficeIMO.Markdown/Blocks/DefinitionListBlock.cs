namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer {
    private readonly List<DefinitionListEntry> _entries = new List<DefinitionListEntry>();

    /// <summary>Typed definition list entries.</summary>
    public IReadOnlyList<DefinitionListEntry> Entries => _entries;

    /// <summary>Legacy tuple-based view over the definition list items.</summary>
    public IList<(string Term, string Definition)> Items { get; }

    /// <summary>Parsed inline representation of the current definition list items.</summary>
    public IReadOnlyList<DefinitionListInlineItem> InlineItems => BuildInlineItems();
    internal List<MarkdownSyntaxNode> SyntaxItems { get; } = new List<MarkdownSyntaxNode>();
    internal MarkdownReaderOptions? ReaderOptions { get; private set; }
    internal MarkdownReaderState? ReaderState { get; private set; }

    /// <summary>Creates a definition list block.</summary>
    public DefinitionListBlock() {
        Items = new LegacyDefinitionListItemList(this);
    }

    internal void SetParsingContext(MarkdownReaderOptions options, MarkdownReaderState state) {
        ReaderOptions = options;
        ReaderState = state;
    }

    /// <summary>Adds a typed definition list entry.</summary>
    public void AddEntry(DefinitionListEntry entry) {
        _entries.Add(entry ?? new DefinitionListEntry());
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _entries.Count; i++) {
            var entry = _entries[i];
            sb.AppendLine(entry.TermMarkdown + ": " + entry.DefinitionMarkdown);
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<dl>");
        for (int index = 0; index < _entries.Count; index++) {
            var entry = _entries[index];
            sb.Append("<dt>");
            sb.Append(entry.Term.RenderHtml());
            sb.Append("</dt>");
            sb.Append("<dd>");
            sb.Append(entry.RenderDefinitionHtml());
            sb.Append("</dd>");
        }
        sb.Append("</dl>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => BuildChildBlocks();

    private IReadOnlyList<DefinitionListInlineItem> BuildInlineItems() {
        if (_entries.Count == 0) {
            return Array.Empty<DefinitionListInlineItem>();
        }

        var options = ReaderOptions ?? new MarkdownReaderOptions();
        var state = ReaderState ?? new MarkdownReaderState();
        var inlineItems = new DefinitionListInlineItem[_entries.Count];

        for (int index = 0; index < _entries.Count; index++) {
            var entry = _entries[index];
            inlineItems[index] = new DefinitionListInlineItem(
                entry.Term ?? new InlineSequence(),
                BuildDefinitionInline(entry, options, state));
        }

        return inlineItems;
    }

    private InlineSequence BuildDefinitionInline(
        DefinitionListEntry entry,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (entry == null || entry.DefinitionBlocks.Count == 0) {
            return new InlineSequence();
        }

        if (entry.DefinitionBlocks.Count == 1 && entry.DefinitionBlocks[0] is ParagraphBlock paragraph) {
            return paragraph.Inlines;
        }

        var markdown = entry.RenderDefinitionMarkdown();
        return string.IsNullOrEmpty(markdown)
            ? new InlineSequence()
            : MarkdownReader.ParseInlineText(markdown, options, state);
    }

    private IReadOnlyList<IMarkdownBlock> BuildChildBlocks() {
        if (_entries.Count == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        var blocks = new List<IMarkdownBlock>();
        for (int i = 0; i < _entries.Count; i++) {
            for (int j = 0; j < _entries[i].DefinitionBlocks.Count; j++) {
                blocks.Add(_entries[i].DefinitionBlocks[j]);
            }
        }
        return blocks;
    }

    private DefinitionListEntry CreateEntryFromLegacyItem(string term, string definition) {
        var options = ReaderOptions ?? new MarkdownReaderOptions();
        var state = ReaderState ?? new MarkdownReaderState();
        var termInlines = string.IsNullOrEmpty(term)
            ? new InlineSequence()
            : MarkdownReader.ParseInlineText(term, options, state);
        var blocks = string.IsNullOrWhiteSpace(definition)
            ? Array.Empty<IMarkdownBlock>()
            : MarkdownReader.ParseBlockFragment(definition, options, state);
        return new DefinitionListEntry(termInlines, blocks);
    }

    private (string Term, string Definition) GetLegacyItem(int index) {
        var entry = _entries[index];
        return (entry.TermMarkdown, entry.DefinitionMarkdown);
    }

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxItems() {
        if (SyntaxItems.Count == _entries.Count && SyntaxItems.Count > 0) {
            return SyntaxItems;
        }

        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var entry in _entries) {
            var term = entry.TermMarkdown;
            var definition = entry.DefinitionMarkdown;
            var definitionChildren = new List<MarkdownSyntaxNode>();
            for (int i = 0; i < entry.DefinitionBlocks.Count; i++) {
                if (entry.DefinitionBlocks[i] is ISyntaxMarkdownBlock syntaxBlock) {
                    definitionChildren.Add(syntaxBlock.BuildSyntaxNode(null));
                } else {
                    definitionChildren.Add(new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.Unknown,
                        literal: entry.DefinitionBlocks[i].RenderMarkdown()));
                }
            }
            if (definitionChildren.Count == 0) {
                definitionChildren.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: definition));
            }

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.DefinitionItem,
                literal: term,
                children: new[] {
                    new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionTerm, literal: term),
                    new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.DefinitionValue,
                        literal: definition,
                        children: definitionChildren)
                }));
        }

        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionList, span, children: BuildSyntaxItems());

    private sealed class LegacyDefinitionListItemList : IList<(string Term, string Definition)> {
        private readonly DefinitionListBlock _owner;

        public LegacyDefinitionListItemList(DefinitionListBlock owner) {
            _owner = owner;
        }

        public int Count => _owner._entries.Count;
        public bool IsReadOnly => false;

        public (string Term, string Definition) this[int index] {
            get => _owner.GetLegacyItem(index);
            set {
                _owner._entries[index] = _owner.CreateEntryFromLegacyItem(value.Term, value.Definition);
                _owner.SyntaxItems.Clear();
            }
        }

        public void Add((string Term, string Definition) item) {
            _owner._entries.Add(_owner.CreateEntryFromLegacyItem(item.Term, item.Definition));
            _owner.SyntaxItems.Clear();
        }

        public void Clear() {
            _owner._entries.Clear();
            _owner.SyntaxItems.Clear();
        }

        public bool Contains((string Term, string Definition) item) => IndexOf(item) >= 0;

        public void CopyTo((string Term, string Definition)[] array, int arrayIndex) {
            for (int i = 0; i < _owner._entries.Count; i++) {
                array[arrayIndex + i] = _owner.GetLegacyItem(i);
            }
        }

        public IEnumerator<(string Term, string Definition)> GetEnumerator() {
            for (int i = 0; i < _owner._entries.Count; i++) {
                yield return _owner.GetLegacyItem(i);
            }
        }

        public int IndexOf((string Term, string Definition) item) {
            for (int i = 0; i < _owner._entries.Count; i++) {
                if (_owner.GetLegacyItem(i).Equals(item)) {
                    return i;
                }
            }
            return -1;
        }

        public void Insert(int index, (string Term, string Definition) item) {
            _owner._entries.Insert(index, _owner.CreateEntryFromLegacyItem(item.Term, item.Definition));
            _owner.SyntaxItems.Clear();
        }

        public bool Remove((string Term, string Definition) item) {
            int index = IndexOf(item);
            if (index < 0) {
                return false;
            }

            RemoveAt(index);
            return true;
        }

        public void RemoveAt(int index) {
            _owner._entries.RemoveAt(index);
            _owner.SyntaxItems.Clear();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
