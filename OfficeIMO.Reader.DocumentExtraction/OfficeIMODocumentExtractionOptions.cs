namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>
/// Settings for <see cref="OfficeIMODocumentExtractionClient"/> and for mapping an OfficeIMO read result.
/// </summary>
public sealed class OfficeIMODocumentExtractionOptions {
    /// <summary>
    /// Reader options applied to every call unless the call supplies
    /// <see cref="OfficeIMOExtractionPropertyNames.ReaderOptionsOverride"/>. Null uses the reader's defaults.
    /// The instance is borrowed, not copied: the reader reads it on each call, so later changes to it affect later
    /// extractions.
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

    /// <summary>
    /// Largest table grid (rows × columns) mapped to a table element. The shared table element allocates its full grid,
    /// so a larger table, such as one with a single very wide row, becomes a text element with role <c>table</c> and
    /// <see cref="OfficeIMOExtractionPropertyNames.TableGridLimitExceeded"/> set. Defaults to 1,000,000 cells.
    /// </summary>
    public long MaxTableCells {
        get => _maxTableCells;
        set => _maxTableCells = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), "The cell limit must be positive.");
    }
    private long _maxTableCells = 1_000_000;

    internal OfficeIMODocumentExtractionOptions Clone() => new OfficeIMODocumentExtractionOptions {
        ReaderOptions = ReaderOptions,
        IncludeImagePayloads = IncludeImagePayloads,
        IncludeRegions = IncludeRegions,
        MaxTableCells = MaxTableCells
    };
}
