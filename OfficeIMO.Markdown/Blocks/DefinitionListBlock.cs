namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock {
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
    /// <summary>Structured definition body blocks flattened across all groups and definitions.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => BuildChildBlocks();
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

    internal void ReplaceEntryTerm(DefinitionListEntry entry, InlineSequence term) {
        if (entry == null) {
            return;
        }

        for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++) {
            var group = _groups[groupIndex];
            for (int termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                if (!entry.IsBoundTo(group, termIndex)) {
                    continue;
                }

                var safeTerm = term ?? new InlineSequence();
                group.ReplaceTerm(termIndex, safeTerm);
                for (int entryIndex = 0; entryIndex < _entries.Count; entryIndex++) {
                    if (_entries[entryIndex].IsBoundTo(group, termIndex)) {
                        _entries[entryIndex].SetTermFromOwner(safeTerm);
                    }
                }

                SyntaxItems.Clear();
                return;
            }
        }

        entry.SetTermFromOwner(term);
    }

    internal void AddParsedEntry(DefinitionListEntry entry, MarkdownSyntaxNode syntaxItem) {
        if (entry == null) {
            throw new ArgumentNullException(nameof(entry));
        }

        if (syntaxItem == null) {
            throw new ArgumentNullException(nameof(syntaxItem));
        }

        AddParsedGroup(
            new DefinitionListGroup(
                new[] { entry.Term },
                new[] { entry.Definition }),
            syntaxItem);
    }

    internal void AddParsedGroup(DefinitionListGroup group, MarkdownSyntaxNode syntaxItem) {
        if (group == null) {
            throw new ArgumentNullException(nameof(group));
        }

        if (syntaxItem == null) {
            throw new ArgumentNullException(nameof(syntaxItem));
        }

        AddGroupCore(group);
        SyntaxItems.Add(syntaxItem);
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
        for (int i = 0; i < _groups.Count; i++) {
            var group = _groups[i];
            var cachedGroup = FindSyntaxGroupForCurrentGroup(group, i);
            if (i > 0) {
                var previousGroup = _groups[i - 1];
                var previousCachedGroup = FindSyntaxGroupForCurrentGroup(previousGroup, i - 1);
                sb.Append(ShouldSeparateGroupsWithBlankLine(previousGroup, previousCachedGroup, group, cachedGroup)
                    ? "\n\n"
                    : "\n");
            }

            AppendGroupMarkdown(sb, group, cachedGroup);
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<dl>");
        for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++) {
            var group = _groups[groupIndex];
            for (int termIndex = 0; termIndex < group.TermItems.Count; termIndex++) {
                var term = group.TermItems[termIndex] ?? new DefinitionListTerm();
                sb.Append("<dt");
                sb.Append(MarkdownHtmlAttributes.Render(term.Attributes, null));
                sb.Append(">");
                sb.Append(term.Inlines.RenderHtml());
                sb.Append(RenderTermGenericAttributeConsumedWhitespace(term));
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

    private static string RenderTermGenericAttributeConsumedWhitespace(DefinitionListTerm term) {
        if (term == null ||
            term.Attributes.IsEmpty ||
            string.IsNullOrEmpty(term.GenericAttributeConsumedWhitespace)) {
            return string.Empty;
        }

        return HtmlTextEncoder.Encode(term.GenericAttributeConsumedWhitespace, HtmlRenderContext.Options);
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxItems;
    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() => BuildSyntaxItems();

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

    private InlineSequence BuildDefinitionInline(
        DefinitionListDefinition definition,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (definition == null || definition.Blocks.Count == 0) {
            return new InlineSequence();
        }

        if (definition.Blocks.Count == 1 && definition.Blocks[0] is ParagraphBlock paragraph) {
            return paragraph.Inlines;
        }

        var markdown = definition.RenderMarkdown();
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

    private static void AppendGroupMarkdown(
        StringBuilder sb,
        DefinitionListGroup? group,
        MarkdownSyntaxNode? cachedGroup) {
        var safeGroup = group ?? new DefinitionListGroup();
        if (ShouldRenderGroupAsMarker(safeGroup, cachedGroup)) {
            AppendMarkerGroupMarkdown(sb, safeGroup);
            return;
        }

        var term = safeGroup.TermItems.Count > 0
            ? safeGroup.TermItems[0]?.Markdown ?? string.Empty
            : string.Empty;
        var definition = safeGroup.Definitions.Count > 0
            ? safeGroup.Definitions[0]
            : null;
        AppendInlineDefinitionMarkdown(sb, term, definition);
    }

    private static bool ShouldRenderGroupAsMarker(DefinitionListGroup group, MarkdownSyntaxNode? cachedGroup) {
        if (group == null) {
            return false;
        }

        if (group.TermItems.Count != 1 || group.Definitions.Count != 1) {
            return true;
        }

        if (cachedGroup == null || cachedGroup.Children.Count < 2) {
            return false;
        }

        var termSpan = cachedGroup.Children
            .FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.DefinitionTerm)
            ?.SourceSpan;
        var definitionSpan = cachedGroup.Children
            .FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.DefinitionValue)
            ?.SourceSpan;

        return termSpan.HasValue &&
               definitionSpan.HasValue &&
               termSpan.Value.EndLine < definitionSpan.Value.StartLine;
    }

    private static bool ShouldSeparateGroupsWithBlankLine(
        DefinitionListGroup previousGroup,
        MarkdownSyntaxNode? previousCachedGroup,
        DefinitionListGroup currentGroup,
        MarkdownSyntaxNode? currentCachedGroup) =>
        ShouldRenderGroupAsMarker(previousGroup, previousCachedGroup) ||
        ShouldRenderGroupAsMarker(currentGroup, currentCachedGroup);

    private static void AppendMarkerGroupMarkdown(StringBuilder sb, DefinitionListGroup group) {
        if (group.TermItems.Count == 0) {
            sb.Append(":   ");
        } else {
            for (int termIndex = 0; termIndex < group.TermItems.Count; termIndex++) {
                if (termIndex > 0) {
                    sb.Append('\n');
                }

                sb.Append(group.TermItems[termIndex]?.Markdown ?? string.Empty);
            }
        }

        if (group.Definitions.Count == 0) {
            sb.Append("\n:   ");
            return;
        }

        for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
            if (definitionIndex > 0 && group.Definitions[definitionIndex - 1]?.ForceParagraphHtml == true) {
                sb.Append("\n\n");
            } else {
                sb.Append('\n');
            }

            AppendMarkerDefinitionMarkdown(sb, group.Definitions[definitionIndex]);
        }
    }

    private static void AppendMarkerDefinitionMarkdown(StringBuilder sb, DefinitionListDefinition? definition) {
        var blocks = definition?.Blocks;
        if (blocks == null || blocks.Count == 0) {
            sb.Append(":   ");
            return;
        }

        if (definition?.HasLeadingBlankLineBeforeBody == true && blocks[0] is ParagraphBlock) {
            sb.Append(":   \n\n");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[0], firstBlock: false, continuationIndent: "    ");
        } else if (blocks[0] is ParagraphBlock) {
            sb.Append(":   ");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[0], firstBlock: true, continuationIndent: "    ");
        } else {
            sb.Append(":   \n");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[0], firstBlock: false, continuationIndent: "    ");
        }

        for (int i = 1; i < blocks.Count; i++) {
            sb.Append(ShouldSeparateDefinitionBlocksWithBlankLine(blocks[i - 1], blocks[i]) ? "\n\n" : "\n");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[i], firstBlock: false, continuationIndent: "    ");
        }
    }

    private static void AppendInlineDefinitionMarkdown(
        StringBuilder sb,
        string termMarkdown,
        DefinitionListDefinition? definition) {
        IReadOnlyList<IMarkdownBlock> blocks = definition == null
            ? Array.Empty<IMarkdownBlock>()
            : definition.Blocks;
        sb.Append(termMarkdown ?? string.Empty).Append(':');
        if (blocks.Count == 0) {
            var definitionMarkdown = definition?.Markdown;
            if (!string.IsNullOrEmpty(definitionMarkdown)) {
                sb.Append(' ').Append(definitionMarkdown);
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
            sb.Append(ShouldSeparateDefinitionBlocksWithBlankLine(blocks[i - 1], blocks[i]) ? "\n\n" : "\n");
            AppendIndentedDefinitionBlockMarkdown(sb, blocks[i], firstBlock: false);
        }
    }

    private static bool ShouldSeparateDefinitionBlocksWithBlankLine(IMarkdownBlock previousBlock, IMarkdownBlock block) =>
        block is ParagraphBlock ||
        (previousBlock is ParagraphBlock && block is HorizontalRuleBlock) ||
        (previousBlock is ParagraphBlock && block is HeadingBlock heading && heading.HasSetextUnderlineMarkerSourceInfo);

    private static void AppendIndentedDefinitionBlockMarkdown(
        StringBuilder sb,
        IMarkdownBlock block,
        bool firstBlock,
        string continuationIndent = "  ") {
        var rendered = RenderDefinitionBlockMarkdown(block)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        var lines = rendered.Split('\n');
        TryGetDefinitionLazyListTailStartLine(block, out var lazyListTailStartLine);
        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) {
                sb.Append('\n');
                if (i < lazyListTailStartLine) {
                    sb.Append(continuationIndent);
                }
            } else if (!firstBlock) {
                sb.Append(continuationIndent);
            }

            sb.Append(i >= lazyListTailStartLine
                ? UnescapeDefinitionLazyListTailLine(lines[i].TrimStart(' '))
                : lines[i]);
        }
    }

    private static string RenderDefinitionBlockMarkdown(IMarkdownBlock? block) {
        if (block is HeadingBlock heading && heading.HasSetextUnderlineMarkerSourceInfo) {
            var marker = !string.IsNullOrWhiteSpace(heading.SetextUnderlineMarkerText)
                ? heading.SetextUnderlineMarkerText!
                : heading.Level == 1
                    ? "==="
                    : "---";
            return heading.Inlines.RenderMarkdown() +
                MarkdownAttributeBlockRenderer.RenderTrailing(heading.Attributes) +
                "\n" +
                marker;
        }

        return block?.RenderMarkdown() ?? string.Empty;
    }

    private static string UnescapeDefinitionLazyListTailLine(string line) =>
        string.IsNullOrEmpty(line)
            ? string.Empty
            : line.Replace("\\|", "|");

    private static bool TryGetDefinitionLazyListTailStartLine(IMarkdownBlock? block, out int lineIndex) {
        lineIndex = int.MaxValue;
        if (block is not IMarkdownListBlock list || list.ListItems.Count == 0) {
            return false;
        }

        var currentLine = 0;
        for (int i = 0; i < list.ListItems.Count; i++) {
            var item = list.ListItems[i];
            if (item == null) {
                currentLine++;
                continue;
            }

            var itemMarkdown = item.RenderMarkdown()
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
            var itemLineCount = Math.Max(1, itemMarkdown.Split('\n').Length);
            if (item.DefinitionLazyParagraphTailContinuation &&
                itemLineCount > 1 &&
                item.AdditionalParagraphs.Count == 0 &&
                item.Children.Count == 0) {
                lineIndex = currentLine + 1;
                return true;
            }

            currentLine += itemLineCount;
        }

        return false;
    }

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxItems() {
        if (SyntaxItems.Count == _groups.Count && SyntaxItems.Count > 0 && SyntaxItemsMatchCurrentGroups()) {
            return SyntaxItems.Select(MarkdownBlockSyntaxBuilder.CloneSyntaxNode).ToArray();
        }

        var nodes = new List<MarkdownSyntaxNode>();
        for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++) {
            var group = _groups[groupIndex];
            var groupChildren = new List<MarkdownSyntaxNode>();
            var cachedGroup = FindSyntaxGroupForCurrentGroup(group, groupIndex);

            for (int termIndex = 0; termIndex < group.TermItems.Count; termIndex++) {
                var termObject = group.TermItems[termIndex] ?? new DefinitionListTerm();
                var term = termObject.Inlines ?? new InlineSequence();
                var cachedTerm = FindCachedDefinitionTerm(cachedGroup, termIndex);
                groupChildren.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    term,
                    cachedTerm?.SourceSpan,
                    literal: termObject.Markdown,
                    associatedObject: termObject));
            }

            for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                var definition = group.Definitions[definitionIndex] ?? new DefinitionListDefinition();
                var definitionLiteral = definition.RenderMarkdown();
                var cachedDefinition = FindCachedDefinitionValue(cachedGroup, group.TermItems.Count + definitionIndex, definition);
                var cachedMarker = FindCachedDefinitionMarker(cachedGroup, definitionIndex, cachedDefinition);
                var definitionChildren = MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(
                    cachedDefinition?.Children,
                    definition.Blocks);

                if (definitionChildren.Count == 0 && !string.IsNullOrEmpty(definitionLiteral)) {
                    var fallbackInlines = BuildDefinitionInline(definition, ReaderOptions ?? new MarkdownReaderOptions(), ReaderState ?? new MarkdownReaderState());
                    definitionChildren = new[] { MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                        MarkdownSyntaxKind.Paragraph,
                        fallbackInlines,
                        literal: definitionLiteral,
                        isGenerated: true) };
                }

                groupChildren.Add(cachedMarker != null
                    ? MarkdownBlockSyntaxBuilder.CloneSyntaxNode(cachedMarker)
                    : CreateGeneratedDefinitionMarker());

                groupChildren.Add(new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionValue,
                    cachedDefinition?.SourceSpan ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionChildren),
                    definitionLiteral,
                    definitionChildren,
                    associatedObject: definition));
            }

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.DefinitionGroup,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(groupChildren),
                children: groupChildren,
                associatedObject: group));
        }

        return nodes;
    }

    private MarkdownSyntaxNode? FindSyntaxGroupForCurrentGroup(DefinitionListGroup group, int preferredIndex) {
        if (group == null || SyntaxItems.Count == 0) {
            return null;
        }

        if (preferredIndex >= 0 &&
            preferredIndex < SyntaxItems.Count &&
            IsSyntaxGroupForCurrentGroup(SyntaxItems[preferredIndex], group)) {
            return SyntaxItems[preferredIndex];
        }

        for (int i = 0; i < SyntaxItems.Count; i++) {
            if (IsSyntaxGroupForCurrentGroup(SyntaxItems[i], group)) {
                return SyntaxItems[i];
            }
        }

        return null;
    }

    private static bool IsSyntaxGroupForCurrentGroup(MarkdownSyntaxNode? syntaxGroup, DefinitionListGroup group) =>
        syntaxGroup != null &&
        syntaxGroup.Kind == MarkdownSyntaxKind.DefinitionGroup &&
        ReferenceEquals(syntaxGroup.AssociatedObject, group);

    private static MarkdownSyntaxNode? FindCachedDefinitionTerm(MarkdownSyntaxNode? cachedGroup, int childIndex) {
        if (cachedGroup == null || childIndex < 0 || childIndex >= cachedGroup.Children.Count) {
            return null;
        }

        var cachedTerm = cachedGroup.Children[childIndex];
        return cachedTerm.Kind == MarkdownSyntaxKind.DefinitionTerm ? cachedTerm : null;
    }

    private static MarkdownSyntaxNode? FindCachedDefinitionValue(
        MarkdownSyntaxNode? cachedGroup,
        int preferredChildIndex,
        DefinitionListDefinition definition) {
        if (cachedGroup == null || definition == null || cachedGroup.Children.Count == 0) {
            return null;
        }

        if (preferredChildIndex >= 0 && preferredChildIndex < cachedGroup.Children.Count) {
            var preferred = cachedGroup.Children[preferredChildIndex];
            if (IsSyntaxDefinitionValueForDefinition(preferred, definition)) {
                return preferred;
            }
        }

        for (int i = 0; i < cachedGroup.Children.Count; i++) {
            if (IsSyntaxDefinitionValueForDefinition(cachedGroup.Children[i], definition)) {
                return cachedGroup.Children[i];
            }
        }

        return null;
    }

    private static MarkdownSyntaxNode? FindCachedDefinitionMarker(
        MarkdownSyntaxNode? cachedGroup,
        int definitionIndex,
        MarkdownSyntaxNode? cachedDefinition) {
        if (cachedGroup == null || definitionIndex < 0 || cachedGroup.Children.Count == 0) {
            return null;
        }

        if (cachedDefinition != null) {
            for (int i = 1; i < cachedGroup.Children.Count; i++) {
                if (ReferenceEquals(cachedGroup.Children[i], cachedDefinition) &&
                    cachedGroup.Children[i - 1].Kind == MarkdownSyntaxKind.DefinitionMarker) {
                    return cachedGroup.Children[i - 1];
                }
            }
        }

        int markerIndex = 0;
        for (int i = 0; i < cachedGroup.Children.Count; i++) {
            if (cachedGroup.Children[i].Kind != MarkdownSyntaxKind.DefinitionMarker) {
                continue;
            }

            if (markerIndex == definitionIndex) {
                return cachedGroup.Children[i];
            }

            markerIndex++;
        }

        return null;
    }

    private static MarkdownSyntaxNode CreateGeneratedDefinitionMarker() =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.DefinitionMarker,
            literal: ":",
            isGenerated: true);

    private static bool IsSyntaxDefinitionValueForDefinition(MarkdownSyntaxNode? syntaxNode, DefinitionListDefinition definition) =>
        syntaxNode != null &&
        syntaxNode.Kind == MarkdownSyntaxKind.DefinitionValue &&
        ReferenceEquals(syntaxNode.AssociatedObject, definition);

    private bool SyntaxItemsMatchCurrentGroups() {
        if (SyntaxItems.Count != _groups.Count) {
            return false;
        }

        for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++) {
            var group = _groups[groupIndex];
            var syntaxGroup = SyntaxItems[groupIndex];
            if (group == null || syntaxGroup == null || syntaxGroup.Kind != MarkdownSyntaxKind.DefinitionGroup) {
                return false;
            }

            if (!ReferenceEquals(syntaxGroup.AssociatedObject, group)) {
                return false;
            }

            var termChildren = syntaxGroup.Children
                .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionTerm)
                .ToArray();
            var markerChildren = syntaxGroup.Children
                .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionMarker)
                .ToArray();
            var definitionChildren = syntaxGroup.Children
                .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionValue)
                .ToArray();
            if (termChildren.Length != group.TermItems.Count ||
                markerChildren.Length != group.Definitions.Count ||
                definitionChildren.Length != group.Definitions.Count) {
                return false;
            }

            for (int termIndex = 0; termIndex < group.TermItems.Count; termIndex++) {
                var termObject = group.TermItems[termIndex] ?? new DefinitionListTerm();
                var term = termObject.Inlines ?? new InlineSequence();
                var syntaxTerm = termChildren[termIndex];
                if (syntaxTerm.Kind != MarkdownSyntaxKind.DefinitionTerm
                    || !ReferenceEquals(syntaxTerm.AssociatedObject, termObject)
                    || !string.Equals(syntaxTerm.Literal ?? string.Empty, term.RenderMarkdown(), StringComparison.Ordinal)) {
                    return false;
                }
            }

            for (int definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                var definition = group.Definitions[definitionIndex] ?? new DefinitionListDefinition();
                var syntaxDefinition = definitionChildren[definitionIndex];
                if (syntaxDefinition.Kind != MarkdownSyntaxKind.DefinitionValue
                    || !ReferenceEquals(syntaxDefinition.AssociatedObject, definition)
                    || !string.Equals(syntaxDefinition.Literal ?? string.Empty, definition.RenderMarkdown(), StringComparison.Ordinal)) {
                    return false;
                }
            }
        }

        return true;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.DefinitionList,
            span,
            children: MarkdownBlockSyntaxBuilder.GetOwnedSyntaxChildrenOrBuild(this),
            associatedObject: this);

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
            for (int termIndex = 0; termIndex < group.TermItems.Count; termIndex++) {
                var term = group.TermItems[termIndex].Inlines;
                var entry = new DefinitionListEntry(term, group.Definitions[definitionIndex]);
                entry.BindToDefinitionList(this, group, termIndex);
                _entries.Add(entry);
            }
        }
    }

    private void RebuildGroupsFromEntries() {
        _groups.Clear();
        for (int i = 0; i < _entries.Count; i++) {
            var entry = _entries[i] ?? new DefinitionListEntry();
            var group = new DefinitionListGroup(
                new[] { entry.Term },
                new[] { entry.Definition });
            _groups.Add(group);
            entry.BindToDefinitionList(this, group, 0);
        }

        SyntaxItems.Clear();
    }
}
