namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Pages creation.</summary>
public sealed class IWorkPagesSaveOptions {
    /// <summary>Gets or sets the maximum number of authored paragraphs, including headers/footers and image/table attachments.</summary>
    public int MaxParagraphs { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined body and header/footer character runs, including attachment markers.</summary>
    public int MaxTextRuns { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum UTF-16 text length across body and header/footer text, image descriptions, table names and cells, including attachment markers and body paragraph terminators.</summary>
    public int MaxTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum combined custom tab stops across body paragraphs.</summary>
    public int MaxTabStops { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum number of decimal lists with authored items.</summary>
    public int MaxNumberedLists { get; set; } = 256;
    /// <summary>Gets or sets the maximum number of inline tables.</summary>
    public int MaxTables { get; set; } = 256;
    /// <summary>Gets or sets the maximum combined dense table grid positions, including unset cells.</summary>
    public int MaxTableCells { get; set; } = 262_144;
    /// <summary>Gets or sets the maximum number of inline image instances.</summary>
    public int MaxImages { get; set; } = 1024;
    /// <summary>Gets or sets the maximum combined encoded image bytes, counting each instance.</summary>
    public long MaxEncodedImageBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Gets or sets the maximum combined RGBA pixel bytes, counting each image instance.</summary>
    public long MaxDecodedImageBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Gets or sets the maximum encoded package and intermediate archive size.</summary>
    public int MaxPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkPagesSaveOptions Snapshot() {
        var copy = new IWorkPagesSaveOptions {
            MaxParagraphs = MaxParagraphs, MaxTextRuns = MaxTextRuns,
            MaxTextCharacters = MaxTextCharacters, MaxTabStops = MaxTabStops,
            MaxNumberedLists = MaxNumberedLists,
            MaxTables = MaxTables, MaxTableCells = MaxTableCells,
            MaxImages = MaxImages, MaxEncodedImageBytes = MaxEncodedImageBytes,
            MaxDecodedImageBytes = MaxDecodedImageBytes,
            MaxPackageBytes = MaxPackageBytes
        };
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxParagraphs, nameof(MaxParagraphs));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxTextRuns, nameof(MaxTextRuns));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxTextCharacters, nameof(MaxTextCharacters));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxTabStops, nameof(MaxTabStops));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxNumberedLists, nameof(MaxNumberedLists));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxTables, nameof(MaxTables));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxTableCells, nameof(MaxTableCells));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxImages, nameof(MaxImages));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxEncodedImageBytes, nameof(MaxEncodedImageBytes));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxDecodedImageBytes, nameof(MaxDecodedImageBytes));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(copy.MaxPackageBytes, nameof(MaxPackageBytes));
        return copy;
    }
}
