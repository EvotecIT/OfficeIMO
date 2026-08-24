namespace OfficeIMO.Latex;

/// <summary>Lossless LaTeX syntax kind.</summary>
public enum LatexSyntaxKind {
    /// <summary>Document root.</summary>
    Document = 0,
    /// <summary>Command with bound argument groups.</summary>
    Command,
    /// <summary>Control sequence token.</summary>
    CommandToken,
    /// <summary>Required brace group.</summary>
    RequiredGroup,
    /// <summary>Optional bracket group.</summary>
    OptionalGroup,
    /// <summary>Group delimiter token.</summary>
    GroupDelimiter,
    /// <summary>Begin/content/end environment.</summary>
    Environment,
    /// <summary>Inline or display math region.</summary>
    Math,
    /// <summary>Math delimiter.</summary>
    MathDelimiter,
    /// <summary>Comment.</summary>
    Comment,
    /// <summary>Whitespace or line ending.</summary>
    Trivia,
    /// <summary>Ordinary token or recoverable unmatched delimiter.</summary>
    Text,
    /// <summary>Opaque inline verbatim command or verbatim-like environment.</summary>
    Verbatim
}

/// <summary>Immutable node in the lossless LaTeX syntax tree.</summary>
public sealed class LatexSyntaxNode {
    private readonly LatexSourceText _source;
    private string? _originalText;
    private int _indexInParent = -1;

    internal LatexSyntaxNode(
        LatexSyntaxKind kind,
        LatexSourceText source,
        LatexSourceSpan span,
        string? value,
        IReadOnlyList<LatexSyntaxNode>? children = null) {
        Kind = kind;
        _source = source;
        Span = span;
        Value = value;
        Children = children ?? Array.Empty<LatexSyntaxNode>();
        for (int index = 0; index < Children.Count; index++) {
            Children[index].Parent = this;
            Children[index]._indexInParent = index;
        }
    }

    /// <summary>Syntax kind.</summary>
    public LatexSyntaxKind Kind { get; }
    /// <summary>Exact source span.</summary>
    public LatexSourceSpan Span { get; }
    /// <summary>Exact source slice.</summary>
    public string OriginalText => _originalText ??= Span.Slice(_source.Text);
    /// <summary>Command or environment name when applicable.</summary>
    public string? Value { get; }
    /// <summary>Parent node.</summary>
    public LatexSyntaxNode? Parent { get; private set; }
    /// <summary>Index in parent or -1.</summary>
    public int IndexInParent => _indexInParent;
    /// <summary>Child nodes in source order.</summary>
    public IReadOnlyList<LatexSyntaxNode> Children { get; }

    /// <summary>Enumerates this node and descendants.</summary>
    public IEnumerable<LatexSyntaxNode> DescendantsAndSelf() {
        LatexSyntaxNode current = this;
        var parents = new Stack<(LatexSyntaxNode Node, int NextChildIndex)>();
        while (true) {
            yield return current;
            if (current.Children.Count > 0) {
                parents.Push((current, 1));
                current = current.Children[0];
                continue;
            }
            while (parents.Count > 0) {
                (LatexSyntaxNode parent, int nextChildIndex) = parents.Pop();
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

    internal bool HasSource(LatexSourceText source) => ReferenceEquals(_source, source);
}

/// <summary>Lossless syntax tree and coverage status.</summary>
public sealed class LatexSyntaxTree {
    internal LatexSyntaxTree(LatexSourceText source, LatexSyntaxNode root) {
        Source = source;
        Root = root;
        IsLossless = Validate(root, source);
    }

    /// <summary>Original source.</summary>
    public LatexSourceText Source { get; }
    /// <summary>Document root.</summary>
    public LatexSyntaxNode Root { get; }
    /// <summary>True when every parent is exactly covered by its children.</summary>
    public bool IsLossless { get; }

    private static bool Validate(LatexSyntaxNode node, LatexSourceText source) {
        if (!node.HasSource(source)) return false;
        if (node.Children.Count == 0) return true;
        int expected = node.Span.Start.Offset;
        for (int index = 0; index < node.Children.Count; index++) {
            LatexSyntaxNode child = node.Children[index];
            if (child.Span.Start.Offset != expected || !Validate(child, source)) return false;
            expected = child.Span.End.Offset;
        }
        return expected == node.Span.End.Offset;
    }
}
