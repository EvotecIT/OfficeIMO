namespace OfficeIMO.Markdown;

/// <summary>
/// Native projection for quoted content.
/// </summary>
public sealed class MarkdownNativeQuoteBlock : MarkdownNativeBlock {
    internal MarkdownNativeQuoteBlock(
        QuoteBlock quote,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children)
        : base(MarkdownNativeBlockKind.Quote, quote, syntaxNode) {
        Quote = quote;
        Lines = quote.Lines;
        MarkerSourceSpans = quote.MarkerSourceSpans;
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        BodySourceSpan = MarkdownNativeContainerSourceSpans.GetAggregateChildSourceSpan(Children);
    }

    /// <summary>Source quote block.</summary>
    public QuoteBlock Quote { get; }

    /// <summary>Raw quote lines captured by the reader when available.</summary>
    public IReadOnlyList<string> Lines { get; }

    /// <summary>Source spans for parsed quote marker tokens.</summary>
    public IReadOnlyList<MarkdownSourceSpan> MarkerSourceSpans { get; }

    /// <summary>Source span for the structured quote body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>Nested native blocks in quote order.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }
}

/// <summary>
/// Native projection for a docs-style callout/admonition.
/// </summary>
public sealed class MarkdownNativeCalloutBlock : MarkdownNativeBlock {
    internal MarkdownNativeCalloutBlock(
        CalloutBlock callout,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children)
        : base(MarkdownNativeBlockKind.Callout, callout, syntaxNode) {
        Callout = callout;
        CalloutKind = callout.Kind;
        OpeningMarkerSourceSpan = FindCalloutChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CalloutOpeningMarker) ?? callout.OpeningMarkerSourceSpan;
        KindSourceSpan = FindCalloutKindSourceSpan(syntaxNode) ?? callout.KindSourceSpan;
        ClosingMarkerSourceSpan = FindCalloutChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CalloutClosingMarker) ?? callout.ClosingMarkerSourceSpan;
        Title = callout.Title;
        TitleSourceSpan = FindCalloutTitleSourceSpan(syntaxNode) ?? callout.TitleSourceSpan;
        TitleInlines = callout.TitleInlines;
        TitleInlineRuns = MarkdownNativeInlineProjection.FromInlineContainerChild(syntaxNode, MarkdownSyntaxKind.CalloutTitle);
        Body = callout.Body;
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        BodySourceSpan = MarkdownNativeContainerSourceSpans.GetAggregateChildSourceSpan(Children);
    }

    /// <summary>Source callout block.</summary>
    public CalloutBlock Callout { get; }

    /// <summary>Callout kind such as info, warning, note, or success.</summary>
    public string CalloutKind { get; }

    /// <summary>Source span for the opening <c>[!</c> marker when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the callout kind token when available.</summary>
    public MarkdownSourceSpan? KindSourceSpan { get; }

    /// <summary>Source span for the closing <c>]</c> marker when available.</summary>
    public MarkdownSourceSpan? ClosingMarkerSourceSpan { get; }

    /// <summary>Plain-text title.</summary>
    public string Title { get; }

    /// <summary>Source span for the explicit callout title when available.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>Structured title inline nodes.</summary>
    public InlineSequence TitleInlines { get; }

    /// <summary>AST-backed native title inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> TitleInlineRuns { get; }

    /// <summary>Rendered markdown body.</summary>
    public string Body { get; }

    /// <summary>Source span for the structured callout body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>Nested native body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    private static MarkdownSourceSpan? FindCalloutChildSourceSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) {
        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == kind) {
                return syntaxNode.Children[i].SourceSpan;
            }
        }

        return null;
    }

    private static MarkdownSourceSpan? FindCalloutKindSourceSpan(MarkdownSyntaxNode syntaxNode) {
        return FindCalloutChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CalloutKind);
    }

    private static MarkdownSourceSpan? FindCalloutTitleSourceSpan(MarkdownSyntaxNode syntaxNode) {
        return FindCalloutChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CalloutTitle);
    }
}

/// <summary>
/// Native projection for a Markdig-style colon-fenced custom container.
/// </summary>
public sealed class MarkdownNativeCustomContainerBlock : MarkdownNativeBlock {
    internal MarkdownNativeCustomContainerBlock(
        CustomContainerBlock container,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children)
        : base(MarkdownNativeBlockKind.CustomContainer, container, syntaxNode) {
        Container = container;
        Name = container.Name;
        Info = container.Info;
        OpeningFence = new string(':', Math.Max(3, container.OpeningFenceLength));
        ClosingFence = new string(':', Math.Max(3, container.ClosingFenceLength));
        OpeningFenceSourceSpan = container.OpeningFenceSourceSpan ?? FindCustomContainerChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CustomContainerOpeningFence);
        InfoSourceSpan = container.InfoSourceSpan ?? FindCustomContainerChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CustomContainerInfo);
        NameSourceSpan = container.NameSourceSpan ?? CustomContainerBlock.CreateNameSourceSpan(InfoSourceSpan, Name);
        ClosingFenceSourceSpan = container.ClosingFenceSourceSpan ?? FindCustomContainerChildSourceSpan(syntaxNode, MarkdownSyntaxKind.CustomContainerClosingFence);
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        BodySourceSpan = MarkdownNativeContainerSourceSpans.GetAggregateChildSourceSpan(Children);
    }

    /// <summary>Source custom container block.</summary>
    public CustomContainerBlock Container { get; }

    /// <summary>First token from the info string, used as the rendered CSS class.</summary>
    public string Name { get; }

    /// <summary>Full source info string after the opening colon fence.</summary>
    public string Info { get; }

    /// <summary>Opening colon fence marker.</summary>
    public string OpeningFence { get; }

    /// <summary>Closing colon fence marker.</summary>
    public string ClosingFence { get; }

    /// <summary>Source span for the opening colon fence marker when available.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; }

    /// <summary>Source span for the container info string when available.</summary>
    public MarkdownSourceSpan? InfoSourceSpan { get; }

    /// <summary>Source span for the first info token used as the rendered CSS class when available.</summary>
    public MarkdownSourceSpan? NameSourceSpan { get; }

    /// <summary>Source span for the structured container body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>Source span for the closing colon fence marker when available.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; }

    /// <summary>Nested native body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    private static MarkdownSourceSpan? FindCustomContainerChildSourceSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) {
        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == kind) {
                return syntaxNode.Children[i].SourceSpan;
            }
        }

        return null;
    }
}

/// <summary>
/// Native projection for a details/disclosure block.
/// </summary>
public sealed class MarkdownNativeDetailsBlock : MarkdownNativeBlock {
    internal MarkdownNativeDetailsBlock(
        DetailsBlock details,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children)
        : base(MarkdownNativeBlockKind.Details, details, syntaxNode) {
        Details = details;
        Open = details.Open;
        OpeningTag = details.OpeningTag;
        ClosingTag = details.ClosingTag;
        OpeningTagSourceSpan = FindDetailsChildSourceSpan(syntaxNode, MarkdownSyntaxKind.DetailsOpeningTag) ?? details.OpeningTagSourceSpan;
        ClosingTagSourceSpan = FindDetailsChildSourceSpan(syntaxNode, MarkdownSyntaxKind.DetailsClosingTag) ?? details.ClosingTagSourceSpan;
        SummaryInlines = details.Summary?.Inlines;
        Summary = SummaryInlines == null ? null : InlinePlainText.Extract(SummaryInlines);
        SummarySourceSpan = FindSummarySourceSpan(syntaxNode) ?? details.Summary?.SourceSpan;
        SummaryOpeningTag = details.Summary?.OpeningTag;
        SummaryText = details.Summary?.SourceText;
        SummaryClosingTag = details.Summary?.ClosingTag;
        SummaryOpeningTagSourceSpan = FindSummaryChildSourceSpan(syntaxNode, MarkdownSyntaxKind.SummaryOpeningTag) ?? details.Summary?.OpeningTagSourceSpan;
        SummaryTextSourceSpan = FindSummaryChildSourceSpan(syntaxNode, MarkdownSyntaxKind.SummaryText) ?? details.Summary?.TextSourceSpan;
        SummaryClosingTagSourceSpan = FindSummaryChildSourceSpan(syntaxNode, MarkdownSyntaxKind.SummaryClosingTag) ?? details.Summary?.ClosingTagSourceSpan;
        SummaryInlineRuns = MarkdownNativeInlineProjection.FromInlineContainerDescendant(syntaxNode, MarkdownSyntaxKind.SummaryText)
            is { Count: > 0 } summaryTextRuns
            ? summaryTextRuns
            : MarkdownNativeInlineProjection.FromInlineContainerChild(syntaxNode, MarkdownSyntaxKind.Summary);
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        BodySourceSpan = MarkdownNativeContainerSourceSpans.GetAggregateChildSourceSpan(Children);
    }

    /// <summary>Source details block.</summary>
    public DetailsBlock Details { get; }

    /// <summary>Whether the details element is initially expanded.</summary>
    public bool Open { get; }

    /// <summary>Exact parsed details opening tag when available.</summary>
    public string? OpeningTag { get; }

    /// <summary>Exact parsed details closing tag when available.</summary>
    public string? ClosingTag { get; }

    /// <summary>Source span for the parsed details opening tag when available.</summary>
    public MarkdownSourceSpan? OpeningTagSourceSpan { get; }

    /// <summary>Source span for the parsed details closing tag when available.</summary>
    public MarkdownSourceSpan? ClosingTagSourceSpan { get; }

    /// <summary>Plain-text summary when available.</summary>
    public string? Summary { get; }

    /// <summary>Source span for the summary element when available.</summary>
    public MarkdownSourceSpan? SummarySourceSpan { get; }

    /// <summary>Exact parsed summary opening tag when available.</summary>
    public string? SummaryOpeningTag { get; }

    /// <summary>Exact parsed summary text when available.</summary>
    public string? SummaryText { get; }

    /// <summary>Exact parsed summary closing tag when available.</summary>
    public string? SummaryClosingTag { get; }

    /// <summary>Source span for the parsed summary opening tag when available.</summary>
    public MarkdownSourceSpan? SummaryOpeningTagSourceSpan { get; }

    /// <summary>Source span for the parsed summary text when available.</summary>
    public MarkdownSourceSpan? SummaryTextSourceSpan { get; }

    /// <summary>Source span for the parsed summary closing tag when available.</summary>
    public MarkdownSourceSpan? SummaryClosingTagSourceSpan { get; }

    /// <summary>Structured summary inline nodes when available.</summary>
    public InlineSequence? SummaryInlines { get; }

    /// <summary>AST-backed native summary inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> SummaryInlineRuns { get; }

    /// <summary>Source span for the structured details body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>Nested native body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    private static MarkdownSourceSpan? FindDetailsChildSourceSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) {
        if (syntaxNode?.Children == null) {
            return null;
        }

        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == kind) {
                return syntaxNode.Children[i].SourceSpan;
            }
        }

        return null;
    }

    private static MarkdownSourceSpan? FindSummarySourceSpan(MarkdownSyntaxNode syntaxNode) {
        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == MarkdownSyntaxKind.Summary) {
                return syntaxNode.Children[i].SourceSpan;
            }
        }

        return null;
    }

    private static MarkdownSourceSpan? FindSummaryChildSourceSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) {
        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            var child = syntaxNode.Children[i];
            if (child.Kind != MarkdownSyntaxKind.Summary) {
                continue;
            }

            for (int j = 0; j < child.Children.Count; j++) {
                if (child.Children[j].Kind == kind) {
                    return child.Children[j].SourceSpan;
                }
            }
        }

        return null;
    }
}

/// <summary>
/// Native projection for a footnote definition.
/// </summary>
public sealed class MarkdownNativeFootnoteDefinitionBlock : MarkdownNativeBlock {
    internal MarkdownNativeFootnoteDefinitionBlock(
        FootnoteDefinitionBlock footnote,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children)
        : base(MarkdownNativeBlockKind.FootnoteDefinition, footnote, syntaxNode) {
        Footnote = footnote;
        Label = footnote.Label;
        OpeningMarkerSourceSpan = FindFootnoteChildSourceSpan(syntaxNode, MarkdownSyntaxKind.FootnoteOpeningMarker) ?? footnote.OpeningMarkerSourceSpan;
        LabelSourceSpan = FindFootnoteLabelSourceSpan(syntaxNode) ?? footnote.LabelSourceSpan;
        SeparatorMarkerSourceSpan = FindFootnoteChildSourceSpan(syntaxNode, MarkdownSyntaxKind.FootnoteSeparatorMarker) ?? footnote.SeparatorMarkerSourceSpan;
        Text = footnote.Text;
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        BodySourceSpan = MarkdownNativeContainerSourceSpans.GetAggregateChildSourceSpan(Children);
    }

    /// <summary>Source footnote definition block.</summary>
    public FootnoteDefinitionBlock Footnote { get; }

    /// <summary>Footnote label without the leading caret marker.</summary>
    public string Label { get; }

    /// <summary>Source span for the opening <c>[^</c> marker when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the footnote label token when available.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; }

    /// <summary>Source span for the <c>]:</c> separator marker when available.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; }

    /// <summary>Rendered markdown text for the definition body.</summary>
    public string Text { get; }

    /// <summary>Source span for the structured definition body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>Nested native definition body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    private static MarkdownSourceSpan? FindFootnoteLabelSourceSpan(MarkdownSyntaxNode syntaxNode) {
        return FindFootnoteChildSourceSpan(syntaxNode, MarkdownSyntaxKind.FootnoteLabel);
    }

    private static MarkdownSourceSpan? FindFootnoteChildSourceSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) {
        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == kind) {
                return syntaxNode.Children[i].SourceSpan;
            }
        }

        return null;
    }

}

/// <summary>
/// Shared source-span helpers for native container body projections.
/// </summary>
internal static class MarkdownNativeContainerSourceSpans {
    /// <summary>Aggregates source spans from native child block syntax nodes when the container body is source-backed.</summary>
    internal static MarkdownSourceSpan? GetAggregateChildSourceSpan(IReadOnlyList<MarkdownNativeBlock> children) {
        if (children == null || children.Count == 0) {
            return null;
        }

        var nodes = new MarkdownSyntaxNode[children.Count];
        for (var i = 0; i < children.Count; i++) {
            nodes[i] = children[i].SyntaxNode;
        }

        return MarkdownBlockSyntaxBuilder.GetAggregateSpan(nodes);
    }
}
