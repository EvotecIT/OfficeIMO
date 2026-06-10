namespace OfficeIMO.Markdown;

/// <summary>
/// Identifies the markdown source represented by a native projection.
/// </summary>
public enum MarkdownNativeDocumentSourceKind {
    /// <summary>The source text is the reader input supplied directly to OfficeIMO.Markdown.</summary>
    ReaderInput,

    /// <summary>The source text is renderer-preprocessed markdown after host preprocessing ran.</summary>
    RendererPreprocessed
}
