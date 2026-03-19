using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Result of parsing markdown through the renderer-owned pipeline.
/// </summary>
public sealed class MarkdownRendererParseResult {
    /// <summary>The final renderer-owned markdown document after parsing and renderer transforms.</summary>
    public MarkdownDoc Document { get; }

    /// <summary>The final markdown text after renderer preprocessing and before parsing.</summary>
    public string PreprocessedMarkdown { get; }

    /// <summary>The original syntax tree produced before document transforms were applied.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }
    /// <summary>The syntax tree corresponding to the final renderer-owned <see cref="Document"/>.</summary>
    public MarkdownSyntaxNode FinalSyntaxTree { get; }

    /// <summary>Document-transform diagnostics from the reader and renderer pipelines.</summary>
    public IReadOnlyList<MarkdownDocumentTransformDiagnostic> TransformDiagnostics { get; }

    /// <summary>Renderer pre-parse processing diagnostics.</summary>
    public IReadOnlyList<MarkdownRendererPreProcessorDiagnostic> PreProcessorDiagnostics { get; }

    internal MarkdownRendererParseResult(
        MarkdownDoc document,
        string preprocessedMarkdown,
        MarkdownSyntaxNode syntaxTree,
        MarkdownSyntaxNode? finalSyntaxTree = null,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null,
        IReadOnlyList<MarkdownRendererPreProcessorDiagnostic>? preProcessorDiagnostics = null) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        PreprocessedMarkdown = preprocessedMarkdown ?? string.Empty;
        SyntaxTree = syntaxTree ?? throw new ArgumentNullException(nameof(syntaxTree));
        FinalSyntaxTree = finalSyntaxTree ?? SyntaxTree;
        TransformDiagnostics = transformDiagnostics ?? Array.Empty<MarkdownDocumentTransformDiagnostic>();
        PreProcessorDiagnostics = preProcessorDiagnostics ?? Array.Empty<MarkdownRendererPreProcessorDiagnostic>();
    }

    /// <summary>Finds the deepest syntax node in the original syntax tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtLine(int lineNumber) => SyntaxTree.FindDeepestNodeAtLine(lineNumber);

    /// <summary>Finds the deepest syntax node in the original syntax tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindDeepestNodeAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the syntax node path from the original document root to the deepest node containing the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtLine(int lineNumber) => SyntaxTree.FindNodePathAtLine(lineNumber);

    /// <summary>Finds the syntax node path from the original document root to the deepest node containing the given 1-based line and column.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindNodePathAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the nearest block-like syntax node in the original syntax tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindNearestBlockAtLine(int lineNumber) => SyntaxTree.FindNearestBlockAtLine(lineNumber);

    /// <summary>Finds the nearest block-like syntax node in the original syntax tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindNearestBlockAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindNearestBlockAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the deepest syntax node in the original syntax tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeContainingSpan(span);

    /// <summary>Finds the syntax node path from the original document root to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathContainingSpan(span);

    /// <summary>Finds the nearest block-like syntax node in the original syntax tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNearestBlockContainingSpan(span);

    /// <summary>Finds the deepest syntax node in the original syntax tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeOverlappingSpan(span);

    /// <summary>Finds the syntax node path from the original document root to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathOverlappingSpan(span);

    /// <summary>Finds the nearest block-like syntax node in the original syntax tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNearestBlockOverlappingSpan(span);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeAtLine(int lineNumber) => FinalSyntaxTree.FindDeepestNodeAtLine(lineNumber);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindDeepestNodeAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the syntax node path from the final document root to the deepest node containing the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathAtLine(int lineNumber) => FinalSyntaxTree.FindNodePathAtLine(lineNumber);

    /// <summary>Finds the syntax node path from the final document root to the deepest node containing the given 1-based line and column.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindNodePathAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockAtLine(int lineNumber) => FinalSyntaxTree.FindNearestBlockAtLine(lineNumber);

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindNearestBlockAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindDeepestNodeContainingSpan(span);

    /// <summary>Finds the syntax node path from the final document root to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNodePathContainingSpan(span);

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNearestBlockContainingSpan(span);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindDeepestNodeOverlappingSpan(span);

    /// <summary>Finds the syntax node path from the final document root to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNodePathOverlappingSpan(span);

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNearestBlockOverlappingSpan(span);
}
