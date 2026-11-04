namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for native Keynote creation.</summary>
public sealed class IWorkKeynoteWriteOptions {
    /// <summary>Gets or sets the maximum number of slides in one output.</summary>
    public int MaxSlides { get; set; } = 4096;
    /// <summary>Gets or sets the maximum number of text boxes across the presentation.</summary>
    public int MaxTextBoxes { get; set; } = 16_384;
    /// <summary>Gets or sets the maximum combined number of retained character runs across all text boxes.</summary>
    public int MaxTextRuns { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined text length in UTF-16 code units across text boxes, table names and table cells, before native paragraph terminators.</summary>
    public int MaxTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum number of image instances across all slides.</summary>
    public int MaxImages { get; set; } = 4096;
    /// <summary>Gets or sets the maximum number of tables across all slides.</summary>
    public int MaxTables { get; set; } = 1024;
    /// <summary>Gets or sets the maximum combined table grid positions, including unset empty cells.</summary>
    public int MaxTableCells { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined encoded image bytes, counting each instance.</summary>
    public long MaxEncodedImageBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Gets or sets the maximum combined RGBA pixel bytes, counting each image instance.</summary>
    public long MaxDecodedImageBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Gets or sets the maximum size of both the encoded package and an intermediate object archive.</summary>
    public int MaxPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkKeynoteWriteOptions Snapshot() {
        var copy = new IWorkKeynoteWriteOptions {
            MaxSlides = MaxSlides, MaxTextBoxes = MaxTextBoxes,
            MaxTextRuns = MaxTextRuns,
            MaxImages = MaxImages, MaxEncodedImageBytes = MaxEncodedImageBytes,
            MaxTables = MaxTables, MaxTableCells = MaxTableCells,
            MaxDecodedImageBytes = MaxDecodedImageBytes,
            MaxTextCharacters = MaxTextCharacters, MaxPackageBytes = MaxPackageBytes
        };
        if (copy.MaxSlides <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSlides));
        if (copy.MaxTextBoxes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextBoxes));
        if (copy.MaxTextRuns <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextRuns));
        if (copy.MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (copy.MaxImages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxImages));
        if (copy.MaxTables <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTables));
        if (copy.MaxTableCells <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTableCells));
        if (copy.MaxEncodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxEncodedImageBytes));
        if (copy.MaxDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedImageBytes));
        if (copy.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
        return copy;
    }
}
