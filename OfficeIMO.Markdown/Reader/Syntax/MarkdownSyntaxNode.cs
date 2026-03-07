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

    /// <summary>Returns this node and all descendant nodes in depth-first order.</summary>
    public IEnumerable<MarkdownSyntaxNode> DescendantsAndSelf() {
        yield return this;
        for (int i = 0; i < Children.Count; i++) {
            foreach (var descendant in Children[i].DescendantsAndSelf()) {
                yield return descendant;
            }
        }
    }

    /// <summary>Finds the deepest node whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtLine(int lineNumber) {
        if (!SourceSpan.HasValue && Children.Count == 0) return null;
        if (SourceSpan.HasValue && !SourceSpan.Value.ContainsLine(lineNumber)) return null;

        for (int i = 0; i < Children.Count; i++) {
            var match = Children[i].FindDeepestNodeAtLine(lineNumber);
            if (match != null) return match;
        }

        return SourceSpan.HasValue ? this : null;
    }

    /// <summary>Finds the node path from this node to the deepest node whose source span contains the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtLine(int lineNumber) {
        var path = new List<MarkdownSyntaxNode>();
        if (!TryBuildNodePathAtLine(lineNumber, path)) return Array.Empty<MarkdownSyntaxNode>();
        return path;
    }

    private bool TryBuildNodePathAtLine(int lineNumber, List<MarkdownSyntaxNode> path) {
        if (!SourceSpan.HasValue && Children.Count == 0) return false;
        if (SourceSpan.HasValue && !SourceSpan.Value.ContainsLine(lineNumber)) return false;

        path.Add(this);
        for (int i = 0; i < Children.Count; i++) {
            if (Children[i].TryBuildNodePathAtLine(lineNumber, path)) return true;
        }

        return SourceSpan.HasValue;
    }
}
