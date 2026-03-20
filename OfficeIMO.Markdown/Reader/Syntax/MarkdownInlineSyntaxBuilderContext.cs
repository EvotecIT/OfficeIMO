namespace OfficeIMO.Markdown;

/// <summary>
/// Helper context passed to <see cref="ISyntaxMarkdownInline"/> implementations when they build
/// custom syntax-tree nodes.
/// </summary>
public sealed class MarkdownInlineSyntaxBuilderContext {
    internal MarkdownInlineSyntaxBuilderContext() {
    }

    /// <summary>
    /// Builds syntax-tree children for the supplied inline sequence.
    /// </summary>
    public IReadOnlyList<MarkdownSyntaxNode> BuildChildren(InlineSequence? sequence) =>
        MarkdownInlineSyntaxBuilder.BuildChildren(sequence);

    /// <summary>
    /// Computes an aggregate source span covering the supplied child nodes when possible.
    /// </summary>
    public MarkdownSourceSpan? GetAggregateSpan(IReadOnlyList<MarkdownSyntaxNode>? children) =>
        MarkdownBlockSyntaxBuilder.GetAggregateSpan(children ?? Array.Empty<MarkdownSyntaxNode>());
}
