namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for template-free native Pages creation.</summary>
public sealed class IWorkPagesWriteOptions {
    /// <summary>Gets or sets the maximum number of body paragraphs.</summary>
    public int MaximumParagraphs { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined retained character runs.</summary>
    public int MaximumTextRuns { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum body length in UTF-16 code units, including paragraph terminators.</summary>
    public int MaximumTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum encoded package and intermediate archive size.</summary>
    public int MaximumPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkPagesWriteOptions Snapshot() {
        var copy = new IWorkPagesWriteOptions {
            MaximumParagraphs = MaximumParagraphs, MaximumTextRuns = MaximumTextRuns,
            MaximumTextCharacters = MaximumTextCharacters, MaximumPackageBytes = MaximumPackageBytes
        };
        if (copy.MaximumParagraphs <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumParagraphs));
        if (copy.MaximumTextRuns <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextRuns));
        if (copy.MaximumTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextCharacters));
        if (copy.MaximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPackageBytes));
        return copy;
    }
}
