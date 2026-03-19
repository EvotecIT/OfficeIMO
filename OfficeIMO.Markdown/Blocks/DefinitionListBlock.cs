namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer {
    private readonly List<DefinitionListGroup> _groups = new List<DefinitionListGroup>();
    private readonly List<DefinitionListEntry> _entries = new List<DefinitionListEntry>();

    /// <summary>Semantic definition-list groups with shared terms and definition bodies.</summary>
    public IReadOnlyList<DefinitionListGroup> Groups => _groups;

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

    internal void ClearSyntaxCache() {
        SyntaxItems.Clear();
    }

    internal void AddParsedEntry(DefinitionListEntry entry, MarkdownSyntaxNode syntaxItem) {
        var safeEntry = entry ?? new DefinitionListEntry();
        AddParsedGroup(
            new DefinitionListGroup(
                new[] { safeEntry.Term },
                new[] { safeEntry.Definition }),
            syntaxItem);
    }

    internal void AddParsedGroup(DefinitionListGroup group, MarkdownSyntaxNode syntaxItem) {
        AddGroupCore(group ?? new DefinitionListGroup());
        SyntaxItems.Add(syntaxItem ?? new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionGroup));
    }

    /// <summary>Adds a typed definition list entry.</summary>
    public void AddEntry(DefinitionListEntry entry) {
        var safeEntry = entry ?? new DefinitionListEntry();
        AddGroup(new DefinitionListGroup(
            new[] { safeEntry.Term },
            new[] { safeEntry.Definition }));
    }

    /// <summary>Adds a semantic definition-list group.</summary>
    public void AddGroup(DefinitionListGroup group) {
        SyntaxItems.Clear();
        AddGroupCore(group ?? new DefinitionListGroup());
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _entries.Count; i++) {
            var entry = _entries[i];
            if (i > 0) {
                sb.Append('\n');
            }

            AppendEntryMarkdown(sb, entry);
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<dl>");
        for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++) {
            var group = _groups[groupIndex];
            for (int termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                sb.Append("<dt>");
                sb.Append(group.Terms[termIndex].RenderHtml());
                sb.Append("</dt>");
            }

            for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                sb.Append("<dd>");
                sb.Append(group.Definitions[definitionIndex].RenderHtml());
                sb.Append("</dd>");
            }
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
        if (_groups.Count == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        var blocks = new List<IMarkdownBlock>();
        for (int i = 0; i < _groups.Count; i++) {
            var definitions = _groups[i].Definitions;
            for (int j = 0; j < definitions.Count; j++) {
                for (int k = 0; k < definitions[j].Blocks.Count; k++) {
                    blocks.Add(definitions[j].Blocks[k]);
                }
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

    private static void AppendEntryMarkdown(StringBuilder sb, DefinitionListEntry? entry) {
        var safeEntry = entry ?? new DefinitionListEntry();
        var blocks = safeEntry.DefinitionBlocks;
        sb.Append(safeEntry.TermMarkdown).Append(':');
        if (blocks.Count == 0) {
            if (!string.IsNullOrEmpty(safeEntry.DefinitionMarkdown)) {
                sb.Append(' ').Append(safeEntry.DefinitionMarkdown);
            }
            return;
        }

        if (blocks[0] is ParagraphBlock) {
            sb.Append(' ');
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[0], firstBlock: true);
        } else {
            sb.Append('\n');
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[0], firstBlock: false);
        }

        for (int i = 1; i < blocks.Count; i++) {
            sb.Append("\n\n");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[i], firstBlock: false);
        }
    }

    private static void AppendIndentedDefinitionBlockMarkdown(StringBuilder sb, IMarkdownBlock block, bool firstBlock) {
        var rendered = (block?.RenderMarkdown() ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        var lines = rendered.Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) {
                sb.Append('\n');
                sb.Append("  ");
            } else if (!firstBlock) {
                sb.Append("  ");
            }

            sb.Append(lines[i]);
        }
    }

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxItems() {
        if (SyntaxItems.Count == _groups.Count && SyntaxItems.Count > 0) {
            return SyntaxItems;
        }

        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var group in _groups) {
            var groupChildren = new List<MarkdownSyntaxNode>();

            for (int termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                var term = group.Terms[termIndex] ?? new InlineSequence();
                groupChildren.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    term,
                    literal: term.RenderMarkdown()));
            }

            for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                var definition = group.Definitions[definitionIndex] ?? new DefinitionListDefinition();
                var definitionLiteral = definition.RenderMarkdown();
                var definitionChildren = new List<MarkdownSyntaxNode>();
                for (int blockIndex = 0; blockIndex < definition.Blocks.Count; blockIndex++) {
                    if (definition.Blocks[blockIndex] != null) {
                        definitionChildren.Add(MarkdownBlockSyntaxBuilder.BuildBlock(definition.Blocks[blockIndex]));
                    }
                }

                if (definitionChildren.Count == 0 && !string.IsNullOrEmpty(definitionLiteral)) {
                    var fallbackEntry = new DefinitionListEntry(new InlineSequence(), definition);
                    var fallbackInlines = BuildDefinitionInline(fallbackEntry, ReaderOptions ?? new MarkdownReaderOptions(), ReaderState ?? new MarkdownReaderState());
                    definitionChildren.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                        MarkdownSyntaxKind.Paragraph,
                        fallbackInlines,
                        literal: definitionLiteral));
                }

                groupChildren.Add(new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionValue,
                    MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionChildren),
                    definitionLiteral,
                    definitionChildren));
            }

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.DefinitionGroup,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(groupChildren),
                children: groupChildren));
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
                _owner.RebuildGroupsFromEntries();
            }
        }

        public void Add((string Term, string Definition) item) {
            _owner._entries.Add(_owner.CreateEntryFromLegacyItem(item.Term, item.Definition));
            _owner.RebuildGroupsFromEntries();
        }

        public void Clear() {
            _owner._entries.Clear();
            _owner.RebuildGroupsFromEntries();
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
            _owner.RebuildGroupsFromEntries();
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
            _owner.RebuildGroupsFromEntries();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private void AddGroupCore(DefinitionListGroup group) {
        _groups.Add(group);

        for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
            for (int termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                var term = group.Terms[termIndex];
                _entries.Add(new DefinitionListEntry(term, group.Definitions[definitionIndex]));
            }
        }
    }

    private void RebuildGroupsFromEntries() {
        _groups.Clear();
        for (int i = 0; i < _entries.Count; i++) {
            var entry = _entries[i] ?? new DefinitionListEntry();
            _groups.Add(new DefinitionListGroup(
                new[] { entry.Term },
                new[] { entry.Definition }));
        }

        SyntaxItems.Clear();
    }
}
