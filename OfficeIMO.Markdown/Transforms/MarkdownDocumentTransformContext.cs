namespace OfficeIMO.Markdown;

/// <summary>
/// Context passed to markdown document transforms.
/// </summary>
public sealed class MarkdownDocumentTransformContext {
    /// <summary>
    /// Source of the current transform pipeline invocation.
    /// </summary>
    public MarkdownDocumentTransformSource Source { get; }

    /// <summary>
    /// Reader options used when the pipeline runs after markdown parsing.
    /// </summary>
    public MarkdownReaderOptions? ReaderOptions { get; }

    /// <summary>
    /// Source options object used by the caller when the pipeline runs outside the markdown reader.
    /// </summary>
    public object? SourceOptions { get; }

    /// <summary>
    /// Optional diagnostics sink populated by the document transform pipeline.
    /// </summary>
    public ICollection<MarkdownDocumentTransformDiagnostic>? Diagnostics { get; }

    /// <summary>
    /// Optional original syntax tree for the document before the current transform pipeline runs.
    /// </summary>
    public MarkdownSyntaxNode? SyntaxTree { get; }

    /// <summary>
    /// Optional per-top-level-block source spans for the document before the current transform pipeline runs.
    /// </summary>
    public IReadOnlyList<MarkdownSourceSpan?>? TopLevelBlockSourceSpans { get; }

    /// <summary>
    /// Normalized markdown source text used to compute <see cref="SyntaxTree"/> source spans.
    /// </summary>
    public string SourceMarkdown { get; }

    /// <summary>
    /// Raw markdown input retained when trivia preservation is enabled; otherwise this falls back to <see cref="SourceMarkdown"/>.
    /// </summary>
    public string OriginalMarkdown { get; }

    /// <summary>
    /// Indicates whether <see cref="OriginalMarkdown"/> contains the exact reader input captured before normalization.
    /// </summary>
    public bool PreservesOriginalMarkdown { get; }

    /// <summary>
    /// Creates a transform context for a reader-driven pipeline.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, MarkdownReaderOptions? readerOptions = null) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceMarkdown = string.Empty;
        OriginalMarkdown = string.Empty;
    }

    /// <summary>
    /// Creates a transform context for an HTML-import pipeline.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, object? sourceOptions) {
        Source = source;
        SourceOptions = sourceOptions;
        SourceMarkdown = string.Empty;
        OriginalMarkdown = string.Empty;
    }

    /// <summary>
    /// Creates a transform context for pipelines that want both reader options and source options,
    /// such as renderer-owned AST transforms that run after markdown parsing.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, MarkdownReaderOptions? readerOptions, object? sourceOptions) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
        SourceMarkdown = string.Empty;
        OriginalMarkdown = string.Empty;
    }

    /// <summary>
    /// Creates a transform context with an explicit diagnostics sink.
    /// </summary>
    public MarkdownDocumentTransformContext(
        MarkdownDocumentTransformSource source,
        MarkdownReaderOptions? readerOptions,
        object? sourceOptions,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
        Diagnostics = diagnostics;
        SourceMarkdown = string.Empty;
        OriginalMarkdown = string.Empty;
    }

    /// <summary>
    /// Creates a transform context with diagnostics and an original syntax tree.
    /// </summary>
    public MarkdownDocumentTransformContext(
        MarkdownDocumentTransformSource source,
        MarkdownReaderOptions? readerOptions,
        object? sourceOptions,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics,
        MarkdownSyntaxNode? syntaxTree) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
        Diagnostics = diagnostics;
        SyntaxTree = syntaxTree;
        SourceMarkdown = string.Empty;
        OriginalMarkdown = string.Empty;
    }

    /// <summary>
    /// Creates a transform context with diagnostics, syntax tree, and explicit top-level block source spans.
    /// </summary>
    public MarkdownDocumentTransformContext(
        MarkdownDocumentTransformSource source,
        MarkdownReaderOptions? readerOptions,
        object? sourceOptions,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics,
        MarkdownSyntaxNode? syntaxTree,
        IReadOnlyList<MarkdownSourceSpan?>? topLevelBlockSourceSpans,
        string? sourceMarkdown = null,
        string? originalMarkdown = null,
        bool preservesOriginalMarkdown = false) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
        Diagnostics = diagnostics;
        SyntaxTree = syntaxTree;
        TopLevelBlockSourceSpans = topLevelBlockSourceSpans;
        SourceMarkdown = sourceMarkdown ?? string.Empty;
        OriginalMarkdown = preservesOriginalMarkdown ? originalMarkdown ?? string.Empty : SourceMarkdown;
        PreservesOriginalMarkdown = preservesOriginalMarkdown;
    }

    /// <summary>
    /// Finds the syntax-tree node associated with a parsed model object before the transform pipeline started.
    /// </summary>
    public MarkdownSyntaxNode? FindSyntaxNode(object associatedObject) {
        if (associatedObject == null || SyntaxTree == null) {
            return null;
        }

        foreach (var node in SyntaxTree.DescendantsAndSelf()) {
            if (ReferenceEquals(node.AssociatedObject, associatedObject)) {
                return node;
            }
        }

        return null;
    }

    /// <summary>
    /// Creates a normalized source slice for the syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        var node = FindSyntaxNode(associatedObject);
        if (node == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(node, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for the supplied syntax node.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        if (syntaxNode == null || !syntaxNode.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(syntaxNode.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) =>
        MarkdownSourceSlice.TryCreate(SourceMarkdown, sourceSpan, MarkdownSourceTextKind.Normalized, out slice);

    /// <summary>
    /// Creates an original-input source slice for the syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(associatedObject, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for the syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        object associatedObject,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        var node = FindSyntaxNode(associatedObject);
        if (node == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.AssociatedObjectNotFound;
            return false;
        }

        return TryCreateOriginalSourceSlice(node, out slice, out failureReason);
    }

    /// <summary>
    /// Creates an original-input source slice for the supplied syntax node.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(syntaxNode, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for the supplied syntax node.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSyntaxNode syntaxNode,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (syntaxNode == null || !syntaxNode.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(syntaxNode.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates an original-input source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(sourceSpan, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan sourceSpan,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        return MarkdownOriginalSourceSliceMapper.TryCreate(
            OriginalMarkdown,
            SourceMarkdown,
            PreservesOriginalMarkdown,
            sourceSpan,
            out slice,
            out failureReason);
    }
}

/// <summary>
/// Known sources for document-transform execution.
/// </summary>
public enum MarkdownDocumentTransformSource {
    /// <summary>Pipeline invoked after markdown parsing.</summary>
    MarkdownReader = 0,
    /// <summary>Pipeline invoked after HTML-to-markdown conversion.</summary>
    HtmlToMarkdown = 1,
    /// <summary>Pipeline invoked by MarkdownRenderer after parsing and before HTML rendering.</summary>
    MarkdownRenderer = 2
}
