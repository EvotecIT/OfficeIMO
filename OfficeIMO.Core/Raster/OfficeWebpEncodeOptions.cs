namespace OfficeIMO.Drawing;

/// <summary>Selects the WebP image compression format.</summary>
public enum OfficeWebpEncodingMode {
    /// <summary>Preserves every RGBA sample using VP8L compression.</summary>
    Lossless,
    /// <summary>Uses VP8 YUV 4:2:0 compression for color and preserves alpha losslessly.</summary>
    Lossy
}

/// <summary>Controls managed WebP encoding.</summary>
public sealed class OfficeWebpEncodeOptions {
    /// <summary>The compression mode. Lossless preserves the existing default encoding contract.</summary>
    public OfficeWebpEncodingMode Mode { get; set; } = OfficeWebpEncodingMode.Lossless;

    /// <summary>VP8 color quality from 1 through 100. This setting applies only to lossy encoding.</summary>
    /// <remarks>Quality 100 still uses lossy YUV 4:2:0 color compression; alpha remains exact.</remarks>
    public int Quality { get; set; } = 85;

    /// <summary>Horizontal resolution written to Exif when physical resolution is enabled.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical resolution written to Exif when physical resolution is enabled.</summary>
    public double DpiY { get; set; } = 96D;

    /// <summary>Writes Exif physical resolution metadata.</summary>
    public bool WritePhysicalResolution { get; set; }

    internal long RetainedManagedBytes { get; set; }
}
