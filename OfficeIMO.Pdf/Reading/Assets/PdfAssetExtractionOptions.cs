namespace OfficeIMO.Pdf;

/// <summary>Controls the files and resource limits of a PDF asset extraction.</summary>
public sealed class PdfAssetExtractionOptions {
    /// <summary>Exports supported image soft masks as straight-alpha, 32-bit RGBA PNGs. When false, uses ordinary image extraction. Default: true.</summary>
    public bool CompositeSoftMasks { get; init; } = true;

    /// <summary>Maximum number of source pages. Default: 500.</summary>
    public int MaximumPages { get; init; } = 500;

    /// <summary>Maximum aggregate uncompressed output file bytes, including reports. Default: 512 MiB.</summary>
    public long MaximumOutputBytes { get; init; } = 512L * 1024L * 1024L;

    /// <summary>Maximum aggregate work units for walking resource graphs and collecting image placements. Default: 16,000,000.</summary>
    public int MaximumInspectionWork { get; init; } = 16_000_000;

    internal void Validate() {
        if (MaximumPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPages));
        if (MaximumOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumOutputBytes));
        if (MaximumInspectionWork <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumInspectionWork));
    }
}
