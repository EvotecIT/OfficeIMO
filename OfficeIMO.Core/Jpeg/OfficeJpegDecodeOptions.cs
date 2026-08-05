namespace OfficeIMO.Drawing;

/// <summary>
/// JPEG decoding options.
/// </summary>
/// <example>
/// <code>
/// var options = new OfficeJpegDecodeOptions(highQualityChroma: true, allowTruncated: true);
/// OfficeRasterImage image = OfficeJpegCodec.Decode(data, options);
/// </code>
/// </example>
public readonly struct OfficeJpegDecodeOptions {
    /// <summary>
    /// Enables higher-quality chroma upsampling when components are subsampled.
    /// </summary>
    public bool HighQualityChroma { get; }

    /// <summary>
    /// Allows truncated scan data (best-effort decode).
    /// </summary>
    public bool AllowTruncated { get; }

    /// <summary>
    /// Creates JPEG decode options.
    /// </summary>
    public OfficeJpegDecodeOptions(bool highQualityChroma = false, bool allowTruncated = false) {
        HighQualityChroma = highQualityChroma;
        AllowTruncated = allowTruncated;
    }
}
