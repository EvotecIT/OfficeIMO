namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Pages creation.</summary>
public sealed class IWorkPagesWriteOptions {
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

    internal IWorkPagesWriteOptions Snapshot() {
        var copy = new IWorkPagesWriteOptions {
            MaxParagraphs = MaxParagraphs, MaxTextRuns = MaxTextRuns,
            MaxTextCharacters = MaxTextCharacters, MaxTabStops = MaxTabStops,
            MaxNumberedLists = MaxNumberedLists,
            MaxTables = MaxTables, MaxTableCells = MaxTableCells,
            MaxImages = MaxImages, MaxEncodedImageBytes = MaxEncodedImageBytes,
            MaxDecodedImageBytes = MaxDecodedImageBytes,
            MaxPackageBytes = MaxPackageBytes
        };
        if (copy.MaxParagraphs <= 0) throw new ArgumentOutOfRangeException(nameof(MaxParagraphs));
        if (copy.MaxTextRuns <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextRuns));
        if (copy.MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (copy.MaxTabStops <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTabStops));
        if (copy.MaxNumberedLists <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNumberedLists));
        if (copy.MaxTables <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTables));
        if (copy.MaxTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTableCells));
        if (copy.MaxImages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxImages));
        if (copy.MaxEncodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxEncodedImageBytes));
        if (copy.MaxDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedImageBytes));
        if (copy.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
        return copy;
    }
}
