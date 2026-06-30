namespace OfficeIMO.Markdown;

/// <summary>
/// Collapsible disclosure block with an optional summary and nested content.
/// </summary>
public sealed class DetailsBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Optional summary displayed in the disclosure header.</summary>
    public SummaryBlock? Summary { get; set; }

    /// <summary>Nested blocks rendered inside the details body.</summary>
    public System.Collections.Generic.List<IMarkdownBlock> Children { get; } = new System.Collections.Generic.List<IMarkdownBlock>();
    /// <summary>Read-only AST-style view of parsed child blocks inside the details body.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => Children;
    /// <summary>Nested syntax nodes captured during parsing, when available.</summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; set; }

    /// <summary>Whether to emit a blank line between the summary and the first child block.</summary>
    public bool InsertBlankLineAfterSummary { get; set; } = true;

    /// <summary>Whether to emit a blank line before the closing tag.</summary>
    public bool InsertBlankLineBeforeClosing { get; set; }

    /// <summary>Whether the details element is initially expanded.</summary>
    public bool Open { get; set; }

    /// <summary>Exact source opening tag for parsed details blocks, when available.</summary>
    internal string? OpeningTag { get; set; }

    /// <summary>Exact source closing tag for parsed details blocks, when available.</summary>
    internal string? ClosingTag { get; set; }

    /// <summary>Source span for the parsed details opening tag, when available.</summary>
    internal MarkdownSourceSpan? OpeningTagSourceSpan { get; set; }

    /// <summary>Source span for the parsed details closing tag, when available.</summary>
    internal MarkdownSourceSpan? ClosingTagSourceSpan { get; set; }

    /// <summary>Creates an empty details block.</summary>
    public DetailsBlock() {
    }

    /// <summary>Creates a details block with a summary and children.</summary>
    public DetailsBlock(SummaryBlock? summary, System.Collections.Generic.IEnumerable<IMarkdownBlock>? children = null, bool open = false) {
        Summary = summary;
        Open = open;
        if (children != null) Children.AddRange(children);
    }

    internal void ClearSyntaxCache() {
        SyntaxChildren = null;
    }

    string IMarkdownBlock.RenderMarkdown() => Render(renderHtmlChildren: false);
    string IMarkdownBlock.RenderHtml() => Render(renderHtmlChildren: true);

    private string Render(bool renderHtmlChildren) {
        var sb = new System.Text.StringBuilder();
        const string NewLine = "\n";
        sb.Append("<details");
        if (Open) sb.Append(" open");
        sb.Append('>');

        if (Summary != null) {
            sb.Append(NewLine);
            sb.Append(renderHtmlChildren
                ? MarkdownBlockRenderDispatcher.RenderHtml(Summary)
                : MarkdownBlockRenderDispatcher.RenderMarkdown(Summary));
        }

        if (Children.Count > 0) {
            sb.Append(NewLine);
            for (int i = 0; i < Children.Count; i++) {
                if (i == 0) {
                    if (Summary != null && InsertBlankLineAfterSummary) sb.Append(NewLine);
                } else {
                    sb.Append(NewLine).Append(NewLine);
                }
                var rendered = renderHtmlChildren
                    ? MarkdownBlockRenderDispatcher.RenderHtml(Children[i])
                    : MarkdownBlockRenderDispatcher.RenderMarkdown(Children[i]);
                sb.Append(rendered);
            }
        }

        sb.Append(NewLine);
        if (InsertBlankLineBeforeClosing && (Children.Count > 0 || Summary != null)) sb.Append(NewLine);
        sb.Append("</details>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        var nodes = new List<MarkdownSyntaxNode>();
        if (OpeningTagSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.DetailsOpeningTag, OpeningTagSourceSpan, OpeningTag));
        }

        if (Summary != null) {
            nodes.Add(MarkdownBlockSyntaxBuilder.BuildBlock(Summary));
        }

        var bodyChildren = MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(SyntaxChildren, ChildBlocks);
        for (int i = 0; i < bodyChildren.Count; i++) {
            nodes.Add(bodyChildren[i]);
        }

        if (ClosingTagSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.DetailsClosingTag, ClosingTagSourceSpan, ClosingTag));
        }

        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Details, span, Open ? "open" : null, ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren(), this);
}

/// <summary>
/// Summary header for a <see cref="DetailsBlock"/>.
/// </summary>
public sealed class SummaryBlock : MarkdownBlock, IMarkdownBlock, IInlineSyntaxMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Inline content inside the &lt;summary&gt; element.</summary>
    public InlineSequence Inlines { get; }
    internal MarkdownSourceSpan? SyntaxSpan { get; set; }

    /// <summary>Create a summary block from an inline sequence.</summary>
    public SummaryBlock(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    /// <summary>Create a summary block with plain text.</summary>
    public SummaryBlock(string? text) {
        Inlines = new InlineSequence().Text(text ?? string.Empty);
    }

    string IMarkdownBlock.RenderMarkdown() => $"<summary>{Inlines.RenderMarkdown()}</summary>";
    string IMarkdownBlock.RenderHtml() => $"<summary>{Inlines.RenderHtml()}</summary>";
    InlineSequence IInlineSyntaxMarkdownBlock.SyntaxInlines => Inlines;
    MarkdownSyntaxKind IInlineSyntaxMarkdownBlock.SyntaxKind => MarkdownSyntaxKind.Summary;
    MarkdownSourceSpan? IInlineSyntaxMarkdownBlock.ProvidedSyntaxSpan => SyntaxSpan;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildInlineBlock(this, span);
}
