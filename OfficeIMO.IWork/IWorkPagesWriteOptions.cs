namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Pages creation.</summary>
public sealed class IWorkPagesWriteOptions {
    /// <summary>Gets or sets the maximum number of authored paragraphs, including headers/footers and image/table attachments.</summary>
    public int MaximumParagraphs { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined body and header/footer character runs, including attachment markers.</summary>
    public int MaximumTextRuns { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum UTF-16 text length across body and header/footer text, image descriptions, table names and cells, including attachment markers and body paragraph terminators.</summary>
    public int MaximumTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum combined custom tab stops across body paragraphs.</summary>
    public int MaximumTabStops { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum number of decimal lists with authored items.</summary>
    public int MaximumNumberedLists { get; set; } = 256;
    /// <summary>Gets or sets the maximum number of inline tables.</summary>
    public int MaximumTables { get; set; } = 256;
    /// <summary>Gets or sets the maximum combined dense table grid positions, including unset cells.</summary>
    public int MaximumTableCells { get; set; } = 262_144;
    /// <summary>Gets or sets the maximum number of inline image instances.</summary>
    public int MaximumImages { get; set; } = 1024;
    /// <summary>Gets or sets the maximum combined encoded image bytes, counting each instance.</summary>
    public long MaximumEncodedImageBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Gets or sets the maximum combined RGBA pixel bytes, counting each image instance.</summary>
    public long MaximumDecodedImageBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Gets or sets the maximum encoded package and intermediate archive size.</summary>
    public int MaximumPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkPagesWriteOptions Snapshot() {
        var copy = new IWorkPagesWriteOptions {
            MaximumParagraphs = MaximumParagraphs, MaximumTextRuns = MaximumTextRuns,
            MaximumTextCharacters = MaximumTextCharacters, MaximumTabStops = MaximumTabStops,
            MaximumNumberedLists = MaximumNumberedLists,
            MaximumTables = MaximumTables, MaximumTableCells = MaximumTableCells,
            MaximumImages = MaximumImages, MaximumEncodedImageBytes = MaximumEncodedImageBytes,
            MaximumDecodedImageBytes = MaximumDecodedImageBytes,
            MaximumPackageBytes = MaximumPackageBytes
        };
        if (copy.MaximumParagraphs <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumParagraphs));
        if (copy.MaximumTextRuns <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextRuns));
        if (copy.MaximumTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextCharacters));
        if (copy.MaximumTabStops <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTabStops));
        if (copy.MaximumNumberedLists <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumNumberedLists));
        if (copy.MaximumTables <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTables));
        if (copy.MaximumTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTableCells));
        if (copy.MaximumImages <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumImages));
        if (copy.MaximumEncodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumEncodedImageBytes));
        if (copy.MaximumDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumDecodedImageBytes));
        if (copy.MaximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPackageBytes));
        return copy;
    }
}
