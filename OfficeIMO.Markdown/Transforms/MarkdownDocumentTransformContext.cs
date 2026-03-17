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
    /// Creates a transform context for a reader-driven pipeline.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, MarkdownReaderOptions? readerOptions = null) {
        Source = source;
        ReaderOptions = readerOptions;
    }

    /// <summary>
    /// Creates a transform context for an HTML-import pipeline.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, object? sourceOptions) {
        Source = source;
        SourceOptions = sourceOptions;
    }

    /// <summary>
    /// Creates a transform context for pipelines that want both reader options and source options,
    /// such as renderer-owned AST transforms that run after markdown parsing.
    /// </summary>
    public MarkdownDocumentTransformContext(MarkdownDocumentTransformSource source, MarkdownReaderOptions? readerOptions, object? sourceOptions) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
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
        IReadOnlyList<MarkdownSourceSpan?>? topLevelBlockSourceSpans) {
        Source = source;
        ReaderOptions = readerOptions;
        SourceOptions = sourceOptions;
        Diagnostics = diagnostics;
        SyntaxTree = syntaxTree;
        TopLevelBlockSourceSpans = topLevelBlockSourceSpans;
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
