namespace OfficeIMO.Markdown;

/// <summary>
/// Markdig-style colon-fenced custom container block rendered as an HTML <c>div</c>.
/// </summary>
public sealed class CustomContainerBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock, ITightListItemHtmlMarkdownBlock {
    private readonly IReadOnlyList<IMarkdownBlock> _childBlocks;
    private MarkdownSourceSpan? _infoSourceSpan;

    /// <summary>First token from the container info string, used as the rendered CSS class.</summary>
    public string Name { get; }

    /// <summary>Full source info string after the opening colon fence.</summary>
    public string Info { get; }

    /// <summary>Number of colon characters in the opening fence.</summary>
    public int OpeningFenceLength { get; }

    /// <summary>Number of colon characters in the closing fence, when the source had one.</summary>
    public int ClosingFenceLength { get; internal set; }

    /// <summary>Parsed child blocks inside the container.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => _childBlocks;

    /// <summary>Nested syntax nodes captured during parsing, when available.</summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; set; }

    /// <summary>Source span for the opening colon fence marker.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; internal set; }

    /// <summary>Source span for the container info string.</summary>
    public MarkdownSourceSpan? InfoSourceSpan {
        get => _infoSourceSpan;
        internal set {
            _infoSourceSpan = value;
            NameSourceSpan = CreateNameSourceSpan(value, Name);
        }
    }

    /// <summary>Source span for the first info token used as the rendered CSS class.</summary>
    public MarkdownSourceSpan? NameSourceSpan { get; internal set; }

    /// <summary>Source span for the closing colon fence marker, when present.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; internal set; }

    /// <summary>Creates a custom container block.</summary>
    public CustomContainerBlock(string? info, IEnumerable<IMarkdownBlock>? children = null, int openingFenceLength = 3) {
        Info = (info ?? string.Empty).Trim();
        Name = GetName(Info);
        OpeningFenceLength = Math.Max(3, openingFenceLength);
        ClosingFenceLength = OpeningFenceLength;
        _childBlocks = CopyChildren(children);
    }

    string IMarkdownBlock.RenderMarkdown() {
        var fenceLength = GetRenderOpeningFenceLength();
        var fence = new string(':', fenceLength);
        var sb = new StringBuilder();
        sb.Append(fence);
        if (!string.IsNullOrWhiteSpace(Info)) {
            sb.Append(' ').Append(Info);
        }

        for (var i = 0; i < ChildBlocks.Count; i++) {
            var rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(ChildBlocks[i]);
            if (string.IsNullOrWhiteSpace(rendered)) {
                continue;
            }

            sb.AppendLine();
            if (i > 0) {
                sb.AppendLine();
            }

            sb.Append(rendered.TrimEnd());
        }

        sb.AppendLine();
        sb.Append(new string(':', Math.Max(fenceLength, Math.Max(3, ClosingFenceLength))));
        return sb.ToString();
    }

    string IMarkdownBlock.RenderHtml() {
        return RenderHtml(tightListItem: false);
    }

    string ITightListItemHtmlMarkdownBlock.RenderTightListItemHtml() {
        return RenderHtml(tightListItem: true);
    }

    private string RenderHtml(bool tightListItem) {
        var sb = new StringBuilder();
        var containerClasses = string.IsNullOrWhiteSpace(Name)
            ? Array.Empty<string>()
            : new[] { Name };
        sb.Append("<div")
            .Append(MarkdownHtmlAttributes.Render(Attributes, HtmlRenderContext.Options, additionalClasses: containerClasses, additionalClassesFirst: true))
            .Append('>');
        for (var i = 0; i < ChildBlocks.Count; i++) {
            if (tightListItem) {
                sb.Append(MarkdownBlockRenderDispatcher.RenderTightListItemHtml(ChildBlocks[i]));
            } else {
                sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(ChildBlocks[i]));
            }
        }

        sb.Append("</div>");
        return sb.ToString();
    }

    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        var children = new List<MarkdownSyntaxNode>();
        if (OpeningFenceSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CustomContainerOpeningFence,
                OpeningFenceSourceSpan,
                new string(':', OpeningFenceLength)));
        }

        if (InfoSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CustomContainerInfo,
                InfoSourceSpan,
                Info));
        }

        var bodyChildren = MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(SyntaxChildren, ChildBlocks);
        for (var i = 0; i < bodyChildren.Count; i++) {
            children.Add(bodyChildren[i]);
        }

        if (ClosingFenceSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CustomContainerClosingFence,
                ClosingFenceSourceSpan,
                new string(':', Math.Max(3, ClosingFenceLength))));
        }

        return children;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CustomContainer,
            span,
            Info,
            ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren(),
            this);

    private static IReadOnlyList<IMarkdownBlock> CopyChildren(IEnumerable<IMarkdownBlock>? children) {
        if (children == null) {
            return Array.Empty<IMarkdownBlock>();
        }

        return children.Where(static child => child != null).ToArray();
    }

    private int GetRenderOpeningFenceLength() {
        var fenceLength = Math.Max(3, OpeningFenceLength);
        var childFenceLength = GetMaxCustomContainerFenceLength(ChildBlocks);
        return childFenceLength >= fenceLength
            ? childFenceLength + 1
            : fenceLength;
    }

    private static int GetMaxCustomContainerFenceLength(IEnumerable<IMarkdownBlock> blocks) {
        var maxFenceLength = 0;
        foreach (var block in blocks) {
            if (block is CustomContainerBlock container) {
                maxFenceLength = Math.Max(maxFenceLength, Math.Max(container.OpeningFenceLength, container.ClosingFenceLength));
            }

            if (block is IMarkdownListBlock listBlock) {
                for (var i = 0; i < listBlock.ListItems.Count; i++) {
                    maxFenceLength = Math.Max(maxFenceLength, GetMaxCustomContainerFenceLength(listBlock.ListItems[i].ChildBlocks));
                }
            }

            if (block is IChildMarkdownBlockContainer childContainer) {
                maxFenceLength = Math.Max(maxFenceLength, GetMaxCustomContainerFenceLength(childContainer.ChildBlocks));
            }
        }

        return maxFenceLength;
    }

    private static string GetName(string info) {
        if (string.IsNullOrWhiteSpace(info)) {
            return string.Empty;
        }

        var trimmed = info.Trim();
        var end = 0;
        while (end < trimmed.Length && !char.IsWhiteSpace(trimmed[end])) {
            end++;
        }

        return trimmed.Substring(0, end);
    }

    internal static MarkdownSourceSpan? CreateNameSourceSpan(MarkdownSourceSpan? infoSourceSpan, string name) {
        if (!infoSourceSpan.HasValue ||
            string.IsNullOrEmpty(name) ||
            !infoSourceSpan.Value.StartColumn.HasValue ||
            !infoSourceSpan.Value.EndColumn.HasValue ||
            infoSourceSpan.Value.StartLine != infoSourceSpan.Value.EndLine) {
            return null;
        }

        var span = infoSourceSpan.Value;
        var endColumn = span.StartColumn.Value + name.Length - 1;
        int? endOffset = null;
        if (span.StartOffset.HasValue) {
            endOffset = span.StartOffset.Value + name.Length - 1;
        }

        return new MarkdownSourceSpan(
            span.StartLine,
            span.StartColumn.Value,
            span.StartLine,
            endColumn,
            span.StartOffset,
            endOffset);
    }
}
