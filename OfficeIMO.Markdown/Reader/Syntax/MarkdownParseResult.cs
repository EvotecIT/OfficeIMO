namespace OfficeIMO.Markdown;

/// <summary>
/// Result of parsing markdown into both the object model and a syntax tree.
/// </summary>
public sealed class MarkdownParseResult {
    /// <summary>The parsed markdown object model.</summary>
    public MarkdownDoc Document { get; }
    /// <summary>
    /// The original syntax tree produced before document transforms were applied.
    /// When a transform replaces the document instance, this tree intentionally drops semantic
    /// <see cref="MarkdownSyntaxNode.AssociatedObject"/> bindings to avoid stale object references.
    /// Use <see cref="FinalSyntaxTree"/> for syntax-to-model navigation against <see cref="Document"/>.
    /// </summary>
    public MarkdownSyntaxNode SyntaxTree { get; }
    /// <summary>The syntax tree corresponding to the final returned <see cref="Document"/>.</summary>
    public MarkdownSyntaxNode FinalSyntaxTree { get; }
    /// <summary>The normalized markdown source text used to compute syntax source spans.</summary>
    public string SourceMarkdown { get; }
    /// <summary>
    /// Raw markdown input retained when <see cref="MarkdownReaderOptions.PreserveTrivia"/> was enabled;
    /// otherwise this falls back to <see cref="SourceMarkdown"/>.
    /// </summary>
    public string OriginalMarkdown { get; }
    /// <summary>
    /// Indicates whether <see cref="OriginalMarkdown"/> contains the exact reader input captured before
    /// input normalization and line-ending normalization.
    /// </summary>
    public bool PreservesOriginalMarkdown { get; }
    /// <summary>Optional document-transform diagnostics captured during parsing.</summary>
    public IReadOnlyList<MarkdownDocumentTransformDiagnostic> TransformDiagnostics { get; }
    /// <summary>Effective reference-style link definitions collected during parsing, in source order where spans are available.</summary>
    public IReadOnlyList<MarkdownReferenceLinkDefinition> ReferenceLinkDefinitions { get; }
    /// <summary>Effective abbreviation definitions collected during parsing, in source order where spans are available.</summary>
    public IReadOnlyList<MarkdownAbbreviationDefinition> AbbreviationDefinitions { get; }

    internal MarkdownParseResult(
        MarkdownDoc document,
        MarkdownSyntaxNode syntaxTree,
        MarkdownSyntaxNode? finalSyntaxTree = null,
        string? sourceMarkdown = null,
        string? originalMarkdown = null,
        bool preservesOriginalMarkdown = false,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null,
        IReadOnlyList<MarkdownReferenceLinkDefinition>? referenceLinkDefinitions = null,
        IReadOnlyList<MarkdownAbbreviationDefinition>? abbreviationDefinitions = null) {
        Document = document;
        SyntaxTree = syntaxTree;
        FinalSyntaxTree = finalSyntaxTree ?? syntaxTree;
        SourceMarkdown = sourceMarkdown ?? string.Empty;
        OriginalMarkdown = preservesOriginalMarkdown ? originalMarkdown ?? string.Empty : SourceMarkdown;
        PreservesOriginalMarkdown = preservesOriginalMarkdown;
        TransformDiagnostics = transformDiagnostics ?? Array.Empty<MarkdownDocumentTransformDiagnostic>();
        ReferenceLinkDefinitions = referenceLinkDefinitions ?? Array.Empty<MarkdownReferenceLinkDefinition>();
        AbbreviationDefinitions = abbreviationDefinitions ?? Array.Empty<MarkdownAbbreviationDefinition>();
        document.AttachParseResult(this);
    }

    /// <summary>
    /// Finds the first node in the final syntax tree associated with the supplied model object.
    /// </summary>
    public MarkdownSyntaxNode? FindFinalNodeForAssociatedObject(object associatedObject) {
        if (associatedObject == null) {
            return null;
        }

        foreach (var node in FinalSyntaxTree.DescendantsAndSelf()) {
            if (ReferenceEquals(node.AssociatedObject, associatedObject)) {
                return node;
            }
        }

        return null;
    }

    /// <summary>
    /// Creates a source slice for the final syntax node associated with the supplied model object.
    /// </summary>
    public bool TryCreateSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        var node = FindFinalNodeForAssociatedObject(associatedObject);
        if (node == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(node, out slice);
    }

    /// <summary>
    /// Creates an original-input source slice for the final syntax node associated with the supplied model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(associatedObject, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for the final syntax node associated with the supplied model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        object associatedObject,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        var node = FindFinalNodeForAssociatedObject(associatedObject);
        if (node == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.AssociatedObjectNotFound;
            return false;
        }

        return TryCreateOriginalSourceSlice(node, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs source spans.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSyntaxNode node, out MarkdownSourceSlice slice) {
        if (node == null || !node.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(node.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs source spans.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan span, out MarkdownSourceSlice slice) =>
        MarkdownSourceSlice.TryCreate(SourceMarkdown, span, MarkdownSourceTextKind.Normalized, out slice);

    /// <summary>
    /// Creates a source slice over the original reader input when it is safely equivalent to the normalized span text.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSyntaxNode node, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(node, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input when it is safely equivalent to the normalized span text.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSyntaxNode node,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (node == null || !node.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(node.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input when it is safely equivalent to the normalized span text.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSourceSpan span, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(span, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input when it is safely equivalent to the normalized span text.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan span,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (!PreservesOriginalMarkdown) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        if (string.Equals(OriginalMarkdown, SourceMarkdown, StringComparison.Ordinal)) {
            if (MarkdownSourceSlice.TryCreate(OriginalMarkdown, span, MarkdownSourceTextKind.Original, out slice)) {
                failureReason = MarkdownOriginalSourceSliceFailureReason.None;
                return true;
            }

            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalSpanUnavailable;
            return false;
        }

        if (!LineEndingsAreEquivalent(OriginalMarkdown, SourceMarkdown)) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent;
            return false;
        }

        if (MarkdownSourceSlice.TryCreateFromLineColumns(OriginalMarkdown, span, MarkdownSourceTextKind.Original, out slice)) {
            failureReason = MarkdownOriginalSourceSliceFailureReason.None;
            return true;
        }

        failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalSpanUnavailable;
        return false;
    }

    /// <summary>Finds the deepest syntax node whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtLine(int lineNumber) => SyntaxTree.FindDeepestNodeAtLine(lineNumber);

    /// <summary>Finds the deepest syntax node whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindDeepestNodeAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the syntax node path from the document root to the deepest node containing the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtLine(int lineNumber) => SyntaxTree.FindNodePathAtLine(lineNumber);

    /// <summary>Finds the syntax node path from the document root to the deepest node containing the given 1-based line and column.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindNodePathAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the nearest block-like syntax node whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindNearestBlockAtLine(int lineNumber) => SyntaxTree.FindNearestBlockAtLine(lineNumber);

    /// <summary>Finds the nearest block-like syntax node whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindNearestBlockAtPosition(int lineNumber, int columnNumber) => SyntaxTree.FindNearestBlockAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the deepest syntax node whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeContainingSpan(span);

    /// <summary>Finds the syntax node path from the document root to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathContainingSpan(span);

    /// <summary>Finds the nearest block-like syntax node whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockContainingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNearestBlockContainingSpan(span);

    /// <summary>Finds the deepest syntax node whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestNodeOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindDeepestNodeOverlappingSpan(span);

    /// <summary>Finds the syntax node path from the document root to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindNodePathOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNodePathOverlappingSpan(span);

    /// <summary>Finds the nearest block-like syntax node whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindNearestBlockOverlappingSpan(MarkdownSourceSpan span) => SyntaxTree.FindNearestBlockOverlappingSpan(span);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeAtLine(int lineNumber) => FinalSyntaxTree.FindDeepestNodeAtLine(lineNumber);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindDeepestNodeAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the syntax node path from the final document root to the deepest node containing the given 1-based line number.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathAtLine(int lineNumber) => FinalSyntaxTree.FindNodePathAtLine(lineNumber);

    /// <summary>Finds the nearest associated object in the final syntax tree at the given 1-based line number.</summary>
    public object? FindFinalAssociatedObjectAtLine(int lineNumber) => FindAssociatedObject(FindFinalNodePathAtLine(lineNumber));

    /// <summary>Finds the nearest associated object of the requested type in the final syntax tree at the given 1-based line number.</summary>
    public TAssociatedObject? FindFinalAssociatedObjectAtLine<TAssociatedObject>(int lineNumber)
        where TAssociatedObject : class =>
        FindAssociatedObject<TAssociatedObject>(FindFinalNodePathAtLine(lineNumber));

    /// <summary>Finds the syntax node path from the final document root to the deepest node containing the given 1-based line and column.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindNodePathAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the nearest associated object in the final syntax tree at the given 1-based line and column.</summary>
    public object? FindFinalAssociatedObjectAtPosition(int lineNumber, int columnNumber) => FindAssociatedObject(FindFinalNodePathAtPosition(lineNumber, columnNumber));

    /// <summary>Finds the nearest associated object of the requested type in the final syntax tree at the given 1-based line and column.</summary>
    public TAssociatedObject? FindFinalAssociatedObjectAtPosition<TAssociatedObject>(int lineNumber, int columnNumber)
        where TAssociatedObject : class =>
        FindAssociatedObject<TAssociatedObject>(FindFinalNodePathAtPosition(lineNumber, columnNumber));

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span contains the given 1-based line number.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockAtLine(int lineNumber) => FinalSyntaxTree.FindNearestBlockAtLine(lineNumber);

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span contains the given 1-based line and column.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockAtPosition(int lineNumber, int columnNumber) => FinalSyntaxTree.FindNearestBlockAtPosition(lineNumber, columnNumber);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindDeepestNodeContainingSpan(span);

    /// <summary>Finds the syntax node path from the final document root to the deepest node whose source span fully contains the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNodePathContainingSpan(span);

    /// <summary>Finds the nearest associated object in the final syntax tree whose source span fully contains the given span.</summary>
    public object? FindFinalAssociatedObjectContainingSpan(MarkdownSourceSpan span) => FindAssociatedObject(FindFinalNodePathContainingSpan(span));

    /// <summary>Finds the nearest associated object of the requested type in the final syntax tree whose source span fully contains the given span.</summary>
    public TAssociatedObject? FindFinalAssociatedObjectContainingSpan<TAssociatedObject>(MarkdownSourceSpan span)
        where TAssociatedObject : class =>
        FindAssociatedObject<TAssociatedObject>(FindFinalNodePathContainingSpan(span));

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span fully contains the given span.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockContainingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNearestBlockContainingSpan(span);

    /// <summary>Finds the deepest syntax node in the final document tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindDeepestFinalNodeOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindDeepestNodeOverlappingSpan(span);

    /// <summary>Finds the syntax node path from the final document root to the deepest node whose source span overlaps the given span.</summary>
    public IReadOnlyList<MarkdownSyntaxNode> FindFinalNodePathOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNodePathOverlappingSpan(span);

    /// <summary>Finds the nearest associated object in the final syntax tree whose source span overlaps the given span.</summary>
    public object? FindFinalAssociatedObjectOverlappingSpan(MarkdownSourceSpan span) => FindAssociatedObject(FindFinalNodePathOverlappingSpan(span));

    /// <summary>Finds the nearest associated object of the requested type in the final syntax tree whose source span overlaps the given span.</summary>
    public TAssociatedObject? FindFinalAssociatedObjectOverlappingSpan<TAssociatedObject>(MarkdownSourceSpan span)
        where TAssociatedObject : class =>
        FindAssociatedObject<TAssociatedObject>(FindFinalNodePathOverlappingSpan(span));

    /// <summary>Finds the nearest block-like syntax node in the final document tree whose source span overlaps the given span.</summary>
    public MarkdownSyntaxNode? FindNearestFinalBlockOverlappingSpan(MarkdownSourceSpan span) => FinalSyntaxTree.FindNearestBlockOverlappingSpan(span);

    private static object? FindAssociatedObject(IReadOnlyList<MarkdownSyntaxNode> path) {
        for (int i = path.Count - 1; i >= 0; i--) {
            if (path[i].AssociatedObject != null) {
                return path[i].AssociatedObject;
            }
        }

        return null;
    }

    private static TAssociatedObject? FindAssociatedObject<TAssociatedObject>(IReadOnlyList<MarkdownSyntaxNode> path)
        where TAssociatedObject : class {
        for (int i = path.Count - 1; i >= 0; i--) {
            if (path[i].AssociatedObject is TAssociatedObject associatedObject) {
                return associatedObject;
            }
        }

        return null;
    }

    private static bool LineEndingsAreEquivalent(string originalMarkdown, string sourceMarkdown) =>
        string.Equals(NormalizeLineEndings(originalMarkdown), sourceMarkdown, StringComparison.Ordinal);

    private static string NormalizeLineEndings(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n');
}
