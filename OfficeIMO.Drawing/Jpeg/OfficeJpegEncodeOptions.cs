namespace OfficeIMO.Drawing;

/// <summary>
/// JPEG encoding options.
/// </summary>
public sealed class OfficeJpegEncodeOptions {
    /// <summary>
    /// JPEG quality (1..100).
    /// </summary>
    public int Quality { get; set; } = 85;

    /// <summary>
    /// Chroma subsampling mode.
    /// </summary>
    public OfficeJpegSubsampling Subsampling { get; set; } = OfficeJpegSubsampling.Y444;

    /// <summary>
    /// Enables progressive JPEG encoding.
    /// </summary>
    public bool Progressive { get; set; }

    /// <summary>
    /// Enables optimized Huffman tables.
    /// </summary>
    public bool OptimizeHuffman { get; set; }

    /// <summary>
    /// Optional metadata segments (EXIF/XMP/ICC).
    /// </summary>
    public OfficeJpegMetadata Metadata { get; set; } = default;

    /// <summary>
    /// Writes a JFIF APP0 header when true.
    /// </summary>
    public bool WriteJfifHeader { get; set; } = true;

    /// <summary>Background used when flattening transparent RGBA pixels into JPEG.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.White;

    /// <summary>Horizontal resolution written to the JFIF header.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical resolution written to the JFIF header.</summary>
    public double DpiY { get; set; } = 96D;
}
