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
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
    }

    /// <summary>Source quote block.</summary>
    public QuoteBlock Quote { get; }

    /// <summary>Raw quote lines captured by the reader when available.</summary>
    public IReadOnlyList<string> Lines { get; }

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
        Title = callout.Title;
        TitleInlines = callout.TitleInlines;
        TitleInlineRuns = MarkdownNativeInlineProjection.FromInlineContainerChild(syntaxNode, MarkdownSyntaxKind.CalloutTitle);
        Body = callout.Body;
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
    }

    /// <summary>Source callout block.</summary>
    public CalloutBlock Callout { get; }

    /// <summary>Callout kind such as info, warning, note, or success.</summary>
    public string CalloutKind { get; }

    /// <summary>Plain-text title.</summary>
    public string Title { get; }

    /// <summary>Structured title inline nodes.</summary>
    public InlineSequence TitleInlines { get; }

    /// <summary>AST-backed native title inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> TitleInlineRuns { get; }

    /// <summary>Rendered markdown body.</summary>
    public string Body { get; }

    /// <summary>Nested native body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }
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
        SummaryInlines = details.Summary?.Inlines;
        Summary = SummaryInlines == null ? null : InlinePlainText.Extract(SummaryInlines);
        SummaryInlineRuns = MarkdownNativeInlineProjection.FromInlineContainerChild(syntaxNode, MarkdownSyntaxKind.Summary);
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
    }

    /// <summary>Source details block.</summary>
    public DetailsBlock Details { get; }

    /// <summary>Whether the details element is initially expanded.</summary>
    public bool Open { get; }

    /// <summary>Plain-text summary when available.</summary>
    public string? Summary { get; }

    /// <summary>Structured summary inline nodes when available.</summary>
    public InlineSequence? SummaryInlines { get; }

    /// <summary>AST-backed native summary inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> SummaryInlineRuns { get; }

    /// <summary>Nested native body blocks.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }
}
