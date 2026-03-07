namespace OfficeIMO.Markdown;

/// <summary>
/// Result of parsing markdown into both the object model and a syntax tree.
/// </summary>
public sealed class MarkdownParseResult {
    /// <summary>The parsed markdown object model.</summary>
    public MarkdownDoc Document { get; }
    /// <summary>The root syntax node for the parsed markdown.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }

    internal MarkdownParseResult(MarkdownDoc document, MarkdownSyntaxNode syntaxTree) {
        Document = document;
        SyntaxTree = syntaxTree;
    }

    /// <summary>Finds the deepest syntax node whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtLine(int lineNumber) => SyntaxTree.FindDeepestNodeAtLine(lineNumber);

    /// <summary>Finds the syntax node path from the document root to the deepest node containing the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtLine(int lineNumber) => SyntaxTree.FindNodePathAtLine(lineNumber);

    /// <summary>Finds the deepest syntax node whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeContainingSpan(span);

    /// <summary>Finds the syntax node path from the document root to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathContainingSpan(span);

    /// <summary>Finds the deepest syntax node whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeOverlappingSpan(span);

    /// <summary>Finds the syntax node path from the document root to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathOverlappingSpan(span);
}
