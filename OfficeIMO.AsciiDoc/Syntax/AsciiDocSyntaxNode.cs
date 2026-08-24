namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Immutable node in the lossless AsciiDoc syntax tree.
/// </summary>
public sealed class AsciiDocSyntaxNode {
    private readonly AsciiDocSourceText _source;
    private string? _originalText;
    private int _indexInParent = -1;

    internal AsciiDocSyntaxNode(
        AsciiDocSyntaxKind kind,
        AsciiDocSourceText source,
        AsciiDocSourceSpan span,
        IReadOnlyList<AsciiDocSyntaxNode>? children = null) {
        Kind = kind;
        _source = source;
        Span = span;
        Children = children ?? Array.Empty<AsciiDocSyntaxNode>();
        for (int index = 0; index < Children.Count; index++) {
            Children[index].Parent = this;
            Children[index]._indexInParent = index;
        }
    }

    /// <summary>Syntax kind.</summary>
    public AsciiDocSyntaxKind Kind { get; }

    /// <summary>Exact half-open source span.</summary>
    public AsciiDocSourceSpan Span { get; }

    /// <summary>Exact source characters covered by this node.</summary>
    public string OriginalText => _originalText ??= Span.Slice(_source.Text);

    /// <summary>Parent node, or null for the document root.</summary>
    public AsciiDocSyntaxNode? Parent { get; private set; }

    /// <summary>Zero-based index within <see cref="Parent"/>, or -1 for the root.</summary>
    public int IndexInParent => _indexInParent;

    /// <summary>Child syntax nodes.</summary>
    public IReadOnlyList<AsciiDocSyntaxNode> Children { get; }

    /// <summary>Enumerates this node and descendants in depth-first order.</summary>
    public IEnumerable<AsciiDocSyntaxNode> DescendantsAndSelf() {
        AsciiDocSyntaxNode current = this;
        var parents = new Stack<(AsciiDocSyntaxNode Node, int NextChildIndex)>();
        while (true) {
            yield return current;
            if (current.Children.Count > 0) {
                parents.Push((current, 1));
                current = current.Children[0];
                continue;
            }
            while (parents.Count > 0) {
                (AsciiDocSyntaxNode parent, int nextChildIndex) = parents.Pop();
                if (nextChildIndex >= parent.Children.Count) continue;
                parents.Push((parent, nextChildIndex + 1));
                current = parent.Children[nextChildIndex];
                goto NextNode;
            }
            yield break;

        NextNode:
            continue;
        }
    }

    internal bool HasSource(AsciiDocSourceText source) => ReferenceEquals(_source, source);

    /// <summary>Finds the deepest node containing a source offset.</summary>
    public AsciiDocSyntaxNode? FindDeepestNodeAtOffset(int offset) {
        bool rootEnd = Parent == null && offset == Span.End.Offset;
        if (!Span.ContainsOffset(offset) && !rootEnd) return null;
        for (int index = 0; index < Children.Count; index++) {
            AsciiDocSyntaxNode? child = Children[index].FindDeepestNodeAtOffset(offset);
            if (child != null) return child;
        }
        return this;
    }
}
