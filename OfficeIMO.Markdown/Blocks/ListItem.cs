namespace OfficeIMO.Markdown;

/// <summary>
/// List item content; supports plain and task (checklist) items.
/// </summary>
public sealed class ListItem : MarkdownObject, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock {
    private readonly ParagraphBlock _leadParagraphBlock;
    private readonly List<ParagraphBlock> _additionalParagraphBlocks = new List<ParagraphBlock>();
    private readonly List<ParagraphBlock> _paragraphBlocks = new List<ParagraphBlock>();
    private readonly List<IMarkdownBlock> _blockChildren = new List<IMarkdownBlock>();

    /// <summary>Inlines representing item content.</summary>
    public InlineSequence Content { get; }
    /// <summary>Additional paragraphs inside the list item (multi-paragraph list items).</summary>
    public List<InlineSequence> AdditionalParagraphs { get; } = new List<InlineSequence>();
    /// <summary>
    /// Paragraph blocks owned by this list item.
    /// This exposes list-item paragraph content as blocks for AST-style consumers.
    /// </summary>
    public IReadOnlyList<ParagraphBlock> ParagraphBlocks {
        get {
            EnsureParagraphBlocks();
            return _paragraphBlocks;
        }
    }
    /// <summary>Nested block content inside the list item (e.g., nested ordered/unordered lists, code blocks).</summary>
    public List<IMarkdownBlock> Children { get; } = new List<IMarkdownBlock>();
    /// <summary>Ordered AST-style view of all list-item child blocks, including lead paragraphs.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks {
        get {
            EnsureBlockChildren();
            return _blockChildren;
        }
    }
    /// <summary>Compatibility alias for <see cref="ChildBlocks"/>.</summary>
    public IReadOnlyList<IMarkdownBlock> BlockChildren => ChildBlocks;
    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;
    /// <summary>True when rendered as a task item (<c>- [ ]</c> or <c>- [x]</c>).</summary>
    public bool IsTask { get; }
    /// <summary>Whether the task is checked.</summary>
    public bool Checked { get; }
    /// <summary>Source span of the list marker token (<c>-</c>, <c>*</c>, <c>+</c>, <c>1.</c>, or <c>1)</c>) when parsed from markdown.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get; internal set; }
    /// <summary>Exact list marker token (<c>-</c>, <c>*</c>, <c>+</c>, <c>1.</c>, or <c>1)</c>) when parsed from markdown.</summary>
    public string? MarkerText { get; internal set; }
    /// <summary>Source span of the task marker token (<c>[ ]</c>, <c>[x]</c>, or <c>[X]</c>) when parsed from markdown.</summary>
    public MarkdownSourceSpan? TaskMarkerSourceSpan { get; internal set; }
    /// <summary>Exact task marker token (<c>[ ]</c>, <c>[x]</c>, or <c>[X]</c>) when parsed from markdown.</summary>
    public string? TaskMarkerText { get; internal set; }
    internal string GenericAttributeConsumedWhitespace { get; set; } = string.Empty;
    internal bool DefinitionLazyParagraphTailContinuation { get; set; }
    /// <summary>Indentation level (0 = top-level). Used for nested lists.</summary>
    public int Level { get; set; }
    /// <summary>Forces paragraph-wrapped loose rendering even when only the first paragraph and child blocks exist.</summary>
    public bool ForceLoose { get; set; }
    internal List<MarkdownSyntaxNode> SyntaxChildren { get; } = new List<MarkdownSyntaxNode>();
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    /// <summary>Creates a plain list item.</summary>
    public ListItem(InlineSequence content) {
        Content = content ?? new InlineSequence();
        _leadParagraphBlock = new ParagraphBlock(Content);
    }

    private ListItem(InlineSequence content, bool isTask, bool isChecked) {
        Content = content ?? new InlineSequence();
        _leadParagraphBlock = new ParagraphBlock(Content);
        IsTask = isTask;
        Checked = isChecked;
    }

    /// <summary>Creates a plain text item.</summary>
    public static ListItem Text(string text) => new ListItem(new InlineSequence().Text(text));
    /// <summary>Creates a link item.</summary>
    public static ListItem Link(string text, string url, string? title = null) => new ListItem(new InlineSequence().Link(text, url, title));
    /// <summary>Creates a task (checklist) item.</summary>
    public static ListItem Task(string text, bool done = false) => new ListItem(new InlineSequence().Text(text), true, done);
    /// <summary>
    /// Creates a task (checklist) item using inline content.
    /// </summary>
    /// <param name="content">Inline content for the list item. When <c>null</c>, an empty sequence is used.</param>
    /// <param name="done">Whether the task should be marked as completed.</param>
    public static ListItem TaskInlines(InlineSequence content, bool done = false) => new ListItem(content ?? new InlineSequence(), true, done);

    internal IEnumerable<InlineSequence> Paragraphs() {
        if (Content.Nodes.Count > 0 || (AdditionalParagraphs.Count == 0 && Children.Count == 0)) {
            yield return Content;
        }
        for (int i = 0; i < AdditionalParagraphs.Count; i++) yield return AdditionalParagraphs[i];
    }

    internal string RenderMarkdown() {
        var parts = Paragraphs().Select(p => p.RenderMarkdown()).ToList();
        if (!Attributes.IsEmpty && parts.Count > 0) {
            var separator = string.IsNullOrEmpty(GenericAttributeConsumedWhitespace)
                ? " "
                : GenericAttributeConsumedWhitespace;
            parts[0] = parts[0].TrimEnd() + separator + MarkdownAttributeBlockRenderer.RenderInlineTrailing(Attributes);
        }

        return string.Join("\n\n", parts);
    }

    internal string RenderHtml() => RenderHtml(forceLoose: false);

    internal string RenderHtml(bool forceLoose, bool renderGenericAttributeConsumedWhitespace = true) {
        bool renderLoose = forceLoose || ForceLoose;
        string checkbox = BuildCheckboxHtml();
        string attributeWhitespace = renderGenericAttributeConsumedWhitespace
            ? RenderGenericAttributeConsumedWhitespace()
            : string.Empty;
        if (!renderLoose && AdditionalParagraphs.Count == 0 && Children.Count == 0) {
            return checkbox + Content.RenderHtml() + attributeWhitespace;
        }

        if (renderLoose
            && Content.Nodes.Count == 0
            && AdditionalParagraphs.Count == 0
            && Children.Count == 0) {
            return checkbox;
        }

        // Tight list behavior: when there is exactly one paragraph, keep it inline even if child blocks exist.
        if (!renderLoose && AdditionalParagraphs.Count == 0) {
            var sbTight = new StringBuilder();
            sbTight.Append(checkbox).Append(Content.RenderHtml()).Append(attributeWhitespace);
            for (int i = 0; i < Children.Count; i++) {
                AppendTightListItemChildSeparator(sbTight, Children[i]);
                sbTight.Append(MarkdownBlockRenderDispatcher.RenderTightListItemHtml(Children[i]));
            }
            return sbTight.ToString();
        }

        // When multiple paragraphs exist, wrap paragraph content in <p> tags.
        var sb = new StringBuilder();
        bool first = true;
        foreach (var p in Paragraphs()) {
            sb.Append("<p>");
            if (first && IsTask) sb.Append(checkbox);
            sb.Append(p.RenderHtml());
            if (first) {
                sb.Append(attributeWhitespace);
            }
            sb.Append("</p>");
            first = false;
        }

        for (int i = 0; i < Children.Count; i++) {
            sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(Children[i]));
        }
        return sb.ToString();
    }

    private static void AppendTightListItemChildSeparator(StringBuilder builder, IMarkdownBlock child) {
        if (child is not TableBlock && child is not CustomContainerBlock ||
            builder.Length == 0 ||
            char.IsWhiteSpace(builder[builder.Length - 1])) {
            return;
        }

        builder.Append(' ');
    }

    private string RenderGenericAttributeConsumedWhitespace() {
        if (string.IsNullOrEmpty(GenericAttributeConsumedWhitespace) || Attributes.IsEmpty) {
            return string.Empty;
        }

        return HtmlTextEncoder.Encode(GenericAttributeConsumedWhitespace, HtmlRenderContext.Options);
    }

    private string BuildCheckboxHtml() {
        if (!IsTask) {
            return string.Empty;
        }

        if (HtmlRenderContext.Options?.GitHubTaskListHtml == true) {
            return Checked
                ? "<input type=\"checkbox\" checked=\"\" disabled=\"\" /> "
                : "<input type=\"checkbox\" disabled=\"\" /> ";
        }

        return "<input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled"
               + (Checked ? " checked" : string.Empty)
               + "> ";
    }

    internal bool TryAbsorbTrailingParagraphBlocks(IReadOnlyList<IMarkdownBlock> trailingBlocks) {
        if (trailingBlocks == null || trailingBlocks.Count == 0) {
            return true;
        }

        for (int i = 0; i < trailingBlocks.Count; i++) {
            if (trailingBlocks[i] is not IParagraphMarkdownBlock paragraph) {
                AdditionalParagraphs.Clear();
                return false;
            }

            AdditionalParagraphs.Add(paragraph.ParagraphInlines);
        }

        return true;
    }

    internal void ReplaceBlockChildren(IReadOnlyList<IMarkdownBlock>? blocks) {
        var incoming = blocks?
            .Where(block => block != null)
            .ToList();
        var preserveSyntaxChildren = HasSameBlockChildren(incoming);
        IMarkdownInline[]? leadInlines = null;
        var additionalParagraphs = new List<InlineSequence>();
        var childBlocks = new List<IMarkdownBlock>();

        if (incoming != null && incoming.Count > 0) {
            var blockIndex = 0;
            if (incoming[0] is ParagraphBlock leadParagraph) {
                leadInlines = leadParagraph.Inlines.Nodes.Where(node => node != null).ToArray();
                blockIndex = 1;
            }

            while (blockIndex < incoming.Count && incoming[blockIndex] is ParagraphBlock additionalParagraph) {
                additionalParagraphs.Add(additionalParagraph.Inlines);
                blockIndex++;
            }

            for (; blockIndex < incoming.Count; blockIndex++) {
                childBlocks.Add(incoming[blockIndex]);
            }
        }

        if (!preserveSyntaxChildren) {
            SyntaxChildren.Clear();
        }

        Content.ReplaceItems(Array.Empty<IMarkdownInline>());
        AdditionalParagraphs.Clear();
        Children.Clear();

        if (incoming == null || incoming.Count == 0) {
            return;
        }

        if (leadInlines != null) {
            Content.ReplaceItems(leadInlines);
        }

        for (int i = 0; i < additionalParagraphs.Count; i++) {
            AdditionalParagraphs.Add(additionalParagraphs[i]);
        }

        for (int i = 0; i < childBlocks.Count; i++) {
            Children.Add(childBlocks[i]);
        }
    }

    private bool HasSameBlockChildren(IReadOnlyList<IMarkdownBlock>? blocks) {
        if (blocks == null) {
            return BlockChildren.Count == 0;
        }

        var current = BlockChildren;
        if (current.Count != blocks.Count) {
            return false;
        }

        for (int i = 0; i < current.Count; i++) {
            if (!ReferenceEquals(current[i], blocks[i])) {
                return false;
            }
        }

        return true;
    }

    internal bool RequiresLooseListRendering() => ForceLoose || AdditionalParagraphs.Count > 0;

    internal MarkdownSyntaxNode BuildSyntaxNode(MarkdownSyntaxNode? nestedList) {
        var children = BuildListMarkerSyntaxNodes();
        children.AddRange(MarkdownBlockSyntaxBuilder.GetOwnedSyntaxChildrenOrBuild(this));
        var attributeNode = MarkdownGenericAttributeSyntaxNodes.Create(this);
        if (attributeNode != null) {
            children.Add(attributeNode);
        }

        if (nestedList != null) {
            children.Add(nestedList);
        }

        string? literal = IsTask
            ? (Checked ? "[x]" : "[ ]")
            : null;

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ListItem,
            MarkdownBlockSyntaxBuilder.GetAggregateSpan(children),
            literal,
            children,
            this,
            attributes: Attributes);
    }

    private List<MarkdownSyntaxNode> BuildListMarkerSyntaxNodes() {
        var children = new List<MarkdownSyntaxNode>(2);
        if (MarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ListMarker, MarkerSourceSpan.Value, MarkerText));
        }

        if (TaskMarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.TaskListMarker, TaskMarkerSourceSpan.Value, TaskMarkerText));
        }

        return children;
    }

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() => BuildOwnedSyntaxChildren();

    private List<MarkdownSyntaxNode> BuildOwnedSyntaxChildren() {
        var blockChildren = BlockChildren;
        if (SyntaxChildren.Count > 0) {
            return BuildCanonicalSyntaxChildrenPreservingSyntaxOnlyNodes(blockChildren);
        }

        return MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(blockChildren).ToList();
    }

    private List<MarkdownSyntaxNode> BuildCanonicalSyntaxChildrenPreservingSyntaxOnlyNodes(IReadOnlyList<IMarkdownBlock> blockChildren) {
        var canonicalChildren = MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(SyntaxChildren, blockChildren).ToList();
        if (!HasSyntaxOnlyDefinitionChildren()) {
            return canonicalChildren;
        }

        var usedCanonicalChildren = new bool[canonicalChildren.Count];
        var children = new List<MarkdownSyntaxNode>(canonicalChildren.Count + SyntaxChildren.Count);
        for (int i = 0; i < SyntaxChildren.Count; i++) {
            var syntaxChild = SyntaxChildren[i];
            if (IsSyntaxOnlyDefinitionChild(syntaxChild)) {
                children.Add(MarkdownBlockSyntaxBuilder.CloneSyntaxNode(syntaxChild));
                continue;
            }

            var canonicalIndex = FindCanonicalChildForCachedSyntax(canonicalChildren, usedCanonicalChildren, syntaxChild);
            if (canonicalIndex >= 0) {
                children.Add(canonicalChildren[canonicalIndex]);
                usedCanonicalChildren[canonicalIndex] = true;
            }
        }

        for (int i = 0; i < canonicalChildren.Count; i++) {
            if (!usedCanonicalChildren[i]) {
                children.Add(canonicalChildren[i]);
            }
        }

        return children;
    }

    private bool HasSyntaxOnlyDefinitionChildren() {
        for (int i = 0; i < SyntaxChildren.Count; i++) {
            if (IsSyntaxOnlyDefinitionChild(SyntaxChildren[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool IsSyntaxOnlyDefinitionChild(MarkdownSyntaxNode node) =>
        node != null
        && (node.Kind == MarkdownSyntaxKind.ReferenceLinkDefinition
            || node.Kind == MarkdownSyntaxKind.AbbreviationDefinition);

    private static int FindCanonicalChildForCachedSyntax(
        IReadOnlyList<MarkdownSyntaxNode> canonicalChildren,
        bool[] usedCanonicalChildren,
        MarkdownSyntaxNode syntaxChild) {
        for (int i = 0; i < canonicalChildren.Count; i++) {
            if (usedCanonicalChildren[i]) {
                continue;
            }

            if (ReferenceEquals(canonicalChildren[i].AssociatedObject, syntaxChild.AssociatedObject)) {
                return i;
            }
        }

        return -1;
    }

    private void EnsureParagraphBlocks() {
        SyncAdditionalParagraphBlocks();

        _paragraphBlocks.Clear();
        if (Content.Nodes.Count > 0 || (AdditionalParagraphs.Count == 0 && Children.Count == 0)) {
            _paragraphBlocks.Add(_leadParagraphBlock);
        }

        for (int i = 0; i < _additionalParagraphBlocks.Count; i++) {
            _paragraphBlocks.Add(_additionalParagraphBlocks[i]);
        }
    }

    private void EnsureBlockChildren() {
        EnsureParagraphBlocks();

        _blockChildren.Clear();
        for (int i = 0; i < _paragraphBlocks.Count; i++) {
            _blockChildren.Add(_paragraphBlocks[i]);
        }

        for (int i = 0; i < Children.Count; i++) {
            _blockChildren.Add(Children[i]);
        }
    }

    private void SyncAdditionalParagraphBlocks() {
        while (_additionalParagraphBlocks.Count > AdditionalParagraphs.Count) {
            _additionalParagraphBlocks.RemoveAt(_additionalParagraphBlocks.Count - 1);
        }

        for (int i = 0; i < AdditionalParagraphs.Count; i++) {
            var paragraph = AdditionalParagraphs[i] ?? new InlineSequence();
            if (i < _additionalParagraphBlocks.Count && ReferenceEquals(_additionalParagraphBlocks[i].Inlines, paragraph)) {
                continue;
            }

            var block = new ParagraphBlock(paragraph);
            if (i < _additionalParagraphBlocks.Count) {
                _additionalParagraphBlocks[i] = block;
            } else {
                _additionalParagraphBlocks.Add(block);
            }
        }
    }

}
