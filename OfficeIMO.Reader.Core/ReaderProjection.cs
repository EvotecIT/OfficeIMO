namespace OfficeIMO.Reader;

/// <summary>Controls which derived representations a supporting format reader constructs.</summary>
public enum ReaderProjection {
    /// <summary>Produces the format reader's ordinary rich document, Markdown and asset projections.</summary>
    Full = 0,
    /// <summary>Produces text and metadata without constructing secondary Markdown, asset, visual, form or OCR projections.
    /// Native parser input and model limits still apply; this is not a hard process-memory limit.</summary>
    TextAndMetadata = 1
}
