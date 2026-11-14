namespace OfficeIMO.Reader;

/// <summary>Controls which derived representations a supporting format reader constructs.</summary>
public enum ReaderProjection {
    /// <summary>Produces the format reader's ordinary rich document, Markdown and asset projections.</summary>
    Full = 0,
    /// <summary>Produces text and metadata without constructing secondary Markdown, asset, visual, form or OCR projections.
    /// Native parser input and model limits still apply; this is not a hard process-memory limit.</summary>
    TextAndMetadata = 1,
    /// <summary>Produces only text chunks whose <c>GetText()</c> output equals the <see cref="TextAndMetadata"/> projection's.
    /// Blocks, tables, pages, links, assets, metadata and nested results are not returned. Handlers that support
    /// <see cref="TextAndMetadata"/> support this projection; a handler may read the text without building its document
    /// model, and otherwise the text-and-metadata result is reduced to its text.</summary>
    Text = 2
}
