namespace OfficeIMO.Markdown;

/// <summary>
/// Helper context passed to <see cref="ISyntaxMarkdownBlockWithContext"/> implementations when they build
/// custom syntax-tree nodes.
/// </summary>
public sealed class MarkdownBlockSyntaxBuilderContext {
    internal MarkdownBlockSyntaxBuilderContext() {
    }

    /// <summary>
    /// Builds a syntax node for a child block using the core reader's standard mapping rules.
    /// </summary>
    public MarkdownSyntaxNode BuildBlock(IMarkdownBlock block, MarkdownSourceSpan? span = null) =>
        MarkdownBlockSyntaxBuilder.BuildBlock(block, span);

    /// <summary>
    /// Builds child syntax nodes for a sequence of child blocks.
    /// </summary>
    public IReadOnlyList<MarkdownSyntaxNode> BuildChildSyntaxNodes(IEnumerable<IMarkdownBlock> children) =>
        MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(children);

    /// <summary>
    /// Builds a syntax node for inline content wrapped in a specific syntax kind.
    /// </summary>
    public MarkdownSyntaxNode BuildInlineContainerNode(
        MarkdownSyntaxKind kind,
        InlineSequence inlines,
        MarkdownSourceSpan? span = null,
        string? literal = null) =>
        MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(kind, inlines, span, literal);

    /// <summary>
    /// Computes an aggregate source span covering the supplied child nodes when possible.
    /// </summary>
    public MarkdownSourceSpan? GetAggregateSpan(IReadOnlyList<MarkdownSyntaxNode>? nodes) =>
        MarkdownBlockSyntaxBuilder.GetAggregateSpan(nodes ?? Array.Empty<MarkdownSyntaxNode>());

    /// <summary>
    /// Normalizes line endings for syntax-node literals to the reader's canonical newline form.
    /// </summary>
    public string NormalizeLiteralLineEndings(string? value) =>
        MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(value);
}
