namespace OfficeIMO.Markdown;

/// <summary>
/// A lightweight syntax-tree node built from the parsed markdown document.
/// </summary>
public sealed class MarkdownSyntaxNode {
    /// <summary>Node kind.</summary>
    public MarkdownSyntaxKind Kind { get; }
    /// <summary>Optional source span from the original markdown.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }
    /// <summary>Optional literal payload for leaf-like nodes.</summary>
    public string? Literal { get; }
    /// <summary>Child syntax nodes.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> Children { get; }

    /// <summary>Create a syntax node.</summary>
    public MarkdownSyntaxNode(MarkdownSyntaxKind kind, MarkdownSourceSpan? sourceSpan = null, string? literal = null, IReadOnlyList<MarkdownSyntaxNode>? children = null) {
        Kind = kind;
        SourceSpan = sourceSpan;
        Literal = literal;
        Children = children ?? Array.Empty<MarkdownSyntaxNode>();
    }
}
