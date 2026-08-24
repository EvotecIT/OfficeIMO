namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Immutable node in the lossless AsciiDoc syntax tree.
/// </summary>
public sealed class AsciiDocSyntaxNode {
    private object _sourceOrOriginalText;
    private readonly int _startOffset;
    private readonly int _endOffset;
    private int _indexInParent = -1;

    internal AsciiDocSyntaxNode(
        AsciiDocSyntaxKind kind,
        AsciiDocSourceText source,
        int startOffset,
        int endOffset,
        IReadOnlyList<AsciiDocSyntaxNode>? children = null) {
        Kind = kind;
        _sourceOrOriginalText = source;
        _startOffset = startOffset;
        _endOffset = endOffset;
        Children = children ?? Array.Empty<AsciiDocSyntaxNode>();
        for (int index = 0; index < Children.Count; index++) {
            Children[index].Parent = this;
            Children[index]._indexInParent = index;
        }
    }

    /// <summary>Syntax kind.</summary>
    public AsciiDocSyntaxKind Kind { get; }

    /// <summary>Exact half-open source span.</summary>
    public AsciiDocSourceSpan Span => GetSource().CreateSpan(_startOffset, _endOffset);

    /// <summary>Exact source characters covered by this node.</summary>
    public string OriginalText {
        get {
            if (_sourceOrOriginalText is string originalText) return originalText;
            var source = (AsciiDocSourceText)_sourceOrOriginalText;
            if (Parent == null) return source.Text;
            originalText = source.Text.Substring(_startOffset, _endOffset - _startOffset);
            _sourceOrOriginalText = originalText;
            return originalText;
        }
    }

    /// <summary>Parent node, or null for the document root.</summary>
    public AsciiDocSyntaxNode? Parent { get; private set; }

    /// <summary>Zero-based index within <see cref="Parent"/>, or -1 for the root.</summary>
    public int IndexInParent => _indexInParent;

    /// <summary>Child syntax nodes.</summary>
    public IReadOnlyList<AsciiDocSyntaxNode> Children { get; }

    internal int StartOffset => _startOffset;

    internal int EndOffset => _endOffset;

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

    internal bool HasSource(AsciiDocSourceText source) => ReferenceEquals(GetSource(), source);

    /// <summary>Finds the deepest node containing a source offset.</summary>
    public AsciiDocSyntaxNode? FindDeepestNodeAtOffset(int offset) {
        bool rootEnd = Parent == null && offset == _endOffset;
        if ((offset < _startOffset || offset >= _endOffset) && !rootEnd) return null;
        for (int index = 0; index < Children.Count; index++) {
            AsciiDocSyntaxNode? child = Children[index].FindDeepestNodeAtOffset(offset);
            if (child != null) return child;
        }
        return this;
    }

    private AsciiDocSourceText GetSource() {
        AsciiDocSyntaxNode? current = this;
        while (current != null) {
            if (current._sourceOrOriginalText is AsciiDocSourceText source) return source;
            current = current.Parent;
        }
        throw new InvalidOperationException("AsciiDoc syntax node is not attached to its source tree.");
    }
}
