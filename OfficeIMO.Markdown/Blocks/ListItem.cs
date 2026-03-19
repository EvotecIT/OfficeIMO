namespace OfficeIMO.Markdown;

/// <summary>
/// List item content; supports plain and task (checklist) items.
/// </summary>
public sealed class ListItem {
    /// <summary>Inlines representing item content.</summary>
    public InlineSequence Content { get; }
    /// <summary>Additional paragraphs inside the list item (multi-paragraph list items).</summary>
    public List<InlineSequence> AdditionalParagraphs { get; } = new List<InlineSequence>();
    /// <summary>
    /// Paragraph blocks owned by this list item.
    /// This exposes list-item paragraph content as blocks for AST-style consumers.
    /// </summary>
    public IReadOnlyList<ParagraphBlock> ParagraphBlocks => BuildParagraphBlocks();
    /// <summary>Nested block content inside the list item (e.g., nested ordered/unordered lists, code blocks).</summary>
    public List<IMarkdownBlock> Children { get; } = new List<IMarkdownBlock>();
    /// <summary>Read-only AST-style view of nested child blocks inside the list item.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => Children;
    /// <summary>Ordered AST-style view of all list-item child blocks, including lead paragraphs.</summary>
    public IReadOnlyList<IMarkdownBlock> BlockChildren => BuildBlockChildren();
    /// <summary>True when rendered as a task item (<c>- [ ]</c> or <c>- [x]</c>).</summary>
    public bool IsTask { get; }
    /// <summary>Whether the task is checked.</summary>
    public bool Checked { get; }
    /// <summary>Indentation level (0 = top-level). Used for nested lists.</summary>
    public int Level { get; set; }
    /// <summary>Forces paragraph-wrapped loose rendering even when only the first paragraph and child blocks exist.</summary>
    public bool ForceLoose { get; set; }
    internal List<MarkdownSyntaxNode> SyntaxChildren { get; } = new List<MarkdownSyntaxNode>();

    /// <summary>Creates a plain list item.</summary>
    public ListItem(InlineSequence content) { Content = content; }
    private ListItem(InlineSequence content, bool isTask, bool isChecked) { Content = content; IsTask = isTask; Checked = isChecked; }

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
        var parts = Paragraphs().Select(p => p.RenderMarkdown());
        return string.Join("\n\n", parts);
    }

    internal string RenderHtml() => RenderHtml(forceLoose: false);

    internal string RenderHtml(bool forceLoose) {
        bool renderLoose = forceLoose || ForceLoose;
        string checkbox = IsTask ? "<input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled" + (Checked ? " checked" : string.Empty) + "> " : string.Empty;
        if (!renderLoose && AdditionalParagraphs.Count == 0 && Children.Count == 0) {
            return checkbox + Content.RenderHtml();
        }

        // Tight list behavior: when there is exactly one paragraph, keep it inline even if child blocks exist.
        if (!renderLoose && AdditionalParagraphs.Count == 0) {
            var sbTight = new StringBuilder();
            sbTight.Append(checkbox).Append(Content.RenderHtml());
            for (int i = 0; i < ChildBlocks.Count; i++) {
                if (ChildBlocks[i] is ITightListItemHtmlMarkdownBlock tightHtmlBlock) {
                    sbTight.Append(tightHtmlBlock.RenderTightListItemHtml());
                } else {
                    sbTight.Append(ChildBlocks[i].RenderHtml());
                }
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
            sb.Append("</p>");
            first = false;
        }

        for (int i = 0; i < ChildBlocks.Count; i++) {
            sb.Append(ChildBlocks[i].RenderHtml());
        }
        return sb.ToString();
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

    internal bool RequiresLooseListRendering() => ForceLoose || AdditionalParagraphs.Count > 0;

    internal MarkdownSyntaxNode BuildSyntaxNode(MarkdownSyntaxNode? nestedList) {
        var children = BuildOwnedSyntaxChildren();

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
            children);
    }

    private List<MarkdownSyntaxNode> BuildOwnedSyntaxChildren() {
        var children = new List<MarkdownSyntaxNode>();
        if (SyntaxChildren.Count > 0) {
            children.AddRange(SyntaxChildren);
            return children;
        }

        var blockChildren = BuildBlockChildren();
        for (int i = 0; i < blockChildren.Count; i++) {
            if (blockChildren[i] is ParagraphBlock paragraph) {
                children.Add(BuildParagraphSyntaxNode(paragraph.Inlines));
            } else {
                children.Add(MarkdownBlockSyntaxBuilder.BuildBlock(blockChildren[i]));
            }
        }

        return children;
    }

    private IReadOnlyList<ParagraphBlock> BuildParagraphBlocks() {
        var paragraphs = new List<ParagraphBlock>();
        if (Content.Nodes.Count > 0 || (AdditionalParagraphs.Count == 0 && Children.Count == 0)) {
            paragraphs.Add(new ParagraphBlock(Content));
        }

        for (int i = 0; i < AdditionalParagraphs.Count; i++) {
            paragraphs.Add(new ParagraphBlock(AdditionalParagraphs[i]));
        }

        return paragraphs;
    }

    private IReadOnlyList<IMarkdownBlock> BuildBlockChildren() {
        var blocks = new List<IMarkdownBlock>();
        var paragraphs = BuildParagraphBlocks();
        for (int i = 0; i < paragraphs.Count; i++) {
            blocks.Add(paragraphs[i]);
        }

        for (int i = 0; i < ChildBlocks.Count; i++) {
            blocks.Add(ChildBlocks[i]);
        }

        return blocks;
    }

    private static MarkdownSyntaxNode BuildParagraphSyntaxNode(InlineSequence paragraph) =>
        MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
            MarkdownSyntaxKind.Paragraph,
            paragraph,
            literal: paragraph.RenderMarkdown());
}
