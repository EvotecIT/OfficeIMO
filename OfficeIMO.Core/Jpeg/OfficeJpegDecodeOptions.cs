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
    /// Allows truncated DCT-based scan data (best-effort decode). Lossless scans require complete entropy data.
    /// </summary>
    public bool AllowTruncated { get; }

    /// <summary>Returns stored sample order without applying the JPEG's optional EXIF display orientation.</summary>
    public bool IgnoreExifOrientation { get; }

    /// <summary>
    /// Creates JPEG decode options.
    /// </summary>
    public OfficeJpegDecodeOptions(bool highQualityChroma = false, bool allowTruncated = false)
        : this(highQualityChroma, allowTruncated, ignoreExifOrientation: false) {
    }

    /// <summary>Creates JPEG decode options with an explicit EXIF orientation policy.</summary>
    public OfficeJpegDecodeOptions(bool highQualityChroma, bool allowTruncated, bool ignoreExifOrientation) {
        HighQualityChroma = highQualityChroma;
        AllowTruncated = allowTruncated;
        IgnoreExifOrientation = ignoreExifOrientation;
    }
}
