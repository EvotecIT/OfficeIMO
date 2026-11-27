namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>
/// Settings for <see cref="OfficeIMODocumentExtractionClient"/> and for mapping an OfficeIMO read result.
/// </summary>
public sealed class OfficeIMODocumentExtractionOptions {
    /// <summary>
    /// Reader options applied to every call unless the call supplies
    /// <see cref="OfficeIMOExtractionPropertyNames.ReaderOptionsOverride"/>. Null uses the reader's defaults.
    /// </summary>
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>
    /// Includes image bytes as <c>DataContent</c> on image elements when the reader kept the payload. Defaults to true.
    /// When false, image elements carry alternative text, OCR text and properties only.
    /// </summary>
    public bool IncludeImagePayloads { get; set; } = true;

    /// <summary>
    /// Adds a source polygon to elements that have page geometry. OfficeIMO reports geometry for PDF, in points with a
    /// bottom-left origin. Defaults to true.
    /// </summary>
    public bool IncludeRegions { get; set; } = true;

    internal OfficeIMODocumentExtractionOptions Clone() => new OfficeIMODocumentExtractionOptions {
        ReaderOptions = ReaderOptions,
        IncludeImagePayloads = IncludeImagePayloads,
        IncludeRegions = IncludeRegions
    };
}
