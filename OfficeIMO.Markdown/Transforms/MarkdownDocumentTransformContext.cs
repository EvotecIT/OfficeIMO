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
}

/// <summary>
/// Known sources for document-transform execution.
/// </summary>
public enum MarkdownDocumentTransformSource {
    /// <summary>Pipeline invoked after markdown parsing.</summary>
    MarkdownReader = 0,
    /// <summary>Pipeline invoked after HTML-to-markdown conversion.</summary>
    HtmlToMarkdown = 1
}
