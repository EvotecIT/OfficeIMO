namespace OfficeIMO.IWork;

/// <summary>Operation-level limits for native Keynote creation.</summary>
public sealed class IWorkKeynoteWriteOptions {
    /// <summary>Gets or sets the maximum number of slides in one output.</summary>
    public int MaximumSlides { get; set; } = 4096;
    /// <summary>Gets or sets the maximum number of text boxes across the presentation.</summary>
    public int MaximumTextBoxes { get; set; } = 16_384;
    /// <summary>Gets or sets the maximum combined number of retained character runs across all text boxes.</summary>
    public int MaximumTextRuns { get; set; } = 65_536;
    /// <summary>Gets or sets the maximum combined text length in UTF-16 code units, before native paragraph terminators.</summary>
    public int MaximumTextCharacters { get; set; } = 8_000_000;
    /// <summary>Gets or sets the maximum number of image instances across all slides.</summary>
    public int MaximumImages { get; set; } = 4096;
    /// <summary>Gets or sets the maximum combined encoded image bytes, counting each instance.</summary>
    public long MaximumEncodedImageBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Gets or sets the maximum combined RGBA pixel bytes, counting each image instance.</summary>
    public long MaximumDecodedImageBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Gets or sets the maximum size of both the encoded package and an intermediate object archive.</summary>
    public int MaximumPackageBytes { get; set; } = 64 * 1024 * 1024;

    internal IWorkKeynoteWriteOptions Snapshot() {
        var copy = new IWorkKeynoteWriteOptions {
            MaximumSlides = MaximumSlides, MaximumTextBoxes = MaximumTextBoxes,
            MaximumTextRuns = MaximumTextRuns,
            MaximumImages = MaximumImages, MaximumEncodedImageBytes = MaximumEncodedImageBytes,
            MaximumDecodedImageBytes = MaximumDecodedImageBytes,
            MaximumTextCharacters = MaximumTextCharacters, MaximumPackageBytes = MaximumPackageBytes
        };
        if (copy.MaximumSlides <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumSlides));
        if (copy.MaximumTextBoxes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextBoxes));
        if (copy.MaximumTextRuns <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextRuns));
        if (copy.MaximumTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumTextCharacters));
        if (copy.MaximumImages <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumImages));
        if (copy.MaximumEncodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumEncodedImageBytes));
        if (copy.MaximumDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumDecodedImageBytes));
        if (copy.MaximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPackageBytes));
        return copy;
    }
}
