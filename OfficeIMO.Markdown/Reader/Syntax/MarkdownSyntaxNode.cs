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

    /// <summary>Finds the deepest node whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeContainingSpan(MarkdownSourceSpan span) {
        if (!SourceSpan.HasValue && Children.Count == 0) return null;
        if (SourceSpan.HasValue && !SourceSpan.Value.Contains(span)) return null;

        for (int i = 0; i < Children.Count; i++) {
            var match = Children[i].FindDeepestNodeContainingSpan(span);
            if (match != null) return match;
        }

        return SourceSpan.HasValue ? this : null;
    }

    /// <summary>Finds the node path from this node to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathContainingSpan(MarkdownSourceSpan span) {
        var path = new List<MarkdownSyntaxNode>();
        if (!TryBuildNodePathContainingSpan(span, path)) return Array.Empty<MarkdownSyntaxNode>();
        return path;
    }

    /// <summary>Finds the deepest node whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeOverlappingSpan(MarkdownSourceSpan span) {
        if (!SourceSpan.HasValue && Children.Count == 0) return null;
        if (SourceSpan.HasValue && !SourceSpan.Value.Overlaps(span)) return null;

        for (int i = 0; i < Children.Count; i++) {
            var match = Children[i].FindDeepestNodeOverlappingSpan(span);
            if (match != null) return match;
        }

        return SourceSpan.HasValue ? this : null;
    }

    /// <summary>Finds the node path from this node to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathOverlappingSpan(MarkdownSourceSpan span) {
        var path = new List<MarkdownSyntaxNode>();
        if (!TryBuildNodePathOverlappingSpan(span, path)) return Array.Empty<MarkdownSyntaxNode>();
        return path;
    }

    /// <summary>Finds the nearest block-like syntax node whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindNearestBlockAtLine(int lineNumber) => FindNearestBlock(FindNodePathAtLine(lineNumber));

    /// <summary>Finds the nearest block-like syntax node whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockContainingSpan(MarkdownSourceSpan span) => FindNearestBlock(FindNodePathContainingSpan(span));

    /// <summary>Finds the nearest block-like syntax node whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockOverlappingSpan(MarkdownSourceSpan span) => FindNearestBlock(FindNodePathOverlappingSpan(span));

    private bool TryBuildNodePathAtLine(int lineNumber, List<MarkdownSyntaxNode> path) {
        if (!SourceSpan.HasValue && Children.Count == 0) return false;
        if (SourceSpan.HasValue && !SourceSpan.Value.ContainsLine(lineNumber)) return false;

        path.Add(this);
        for (int i = 0; i < Children.Count; i++) {
            if (Children[i].TryBuildNodePathAtLine(lineNumber, path)) return true;
        }

        return SourceSpan.HasValue;
    }

    private bool TryBuildNodePathContainingSpan(MarkdownSourceSpan span, List<MarkdownSyntaxNode> path) {
        if (!SourceSpan.HasValue && Children.Count == 0) return false;
        if (SourceSpan.HasValue && !SourceSpan.Value.Contains(span)) return false;

        path.Add(this);
        for (int i = 0; i < Children.Count; i++) {
            if (Children[i].TryBuildNodePathContainingSpan(span, path)) return true;
        }

        return SourceSpan.HasValue;
    }

    private bool TryBuildNodePathOverlappingSpan(MarkdownSourceSpan span, List<MarkdownSyntaxNode> path) {
        if (!SourceSpan.HasValue && Children.Count == 0) return false;
        if (SourceSpan.HasValue && !SourceSpan.Value.Overlaps(span)) return false;

        path.Add(this);
        for (int i = 0; i < Children.Count; i++) {
            if (Children[i].TryBuildNodePathOverlappingSpan(span, path)) return true;
        }

        return SourceSpan.HasValue;
    }

    private static MarkdownSyntaxNode? FindNearestBlock(IReadOnlyList<MarkdownSyntaxNode> path) {
        for (int i = path.Count - 1; i >= 0; i--) {
            if (IsBlockLikeKind(path[i].Kind)) return path[i];
        }

        return null;
    }

    private static bool IsBlockLikeKind(MarkdownSyntaxKind kind) {
        switch (kind) {
            case MarkdownSyntaxKind.Document:
            case MarkdownSyntaxKind.Heading:
            case MarkdownSyntaxKind.Paragraph:
            case MarkdownSyntaxKind.Quote:
            case MarkdownSyntaxKind.UnorderedList:
            case MarkdownSyntaxKind.OrderedList:
            case MarkdownSyntaxKind.ListItem:
            case MarkdownSyntaxKind.CodeBlock:
            case MarkdownSyntaxKind.Table:
            case MarkdownSyntaxKind.TableHeader:
            case MarkdownSyntaxKind.TableRow:
            case MarkdownSyntaxKind.HorizontalRule:
            case MarkdownSyntaxKind.Image:
            case MarkdownSyntaxKind.Callout:
            case MarkdownSyntaxKind.DefinitionList:
            case MarkdownSyntaxKind.DefinitionItem:
            case MarkdownSyntaxKind.FootnoteDefinition:
            case MarkdownSyntaxKind.Details:
            case MarkdownSyntaxKind.Summary:
            case MarkdownSyntaxKind.FrontMatter:
            case MarkdownSyntaxKind.HtmlRaw:
            case MarkdownSyntaxKind.HtmlComment:
            case MarkdownSyntaxKind.Toc:
            case MarkdownSyntaxKind.TocPlaceholder:
            case MarkdownSyntaxKind.Unknown:
                return true;
            default:
                return false;
        }
    }
}
