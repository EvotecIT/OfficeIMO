namespace OfficeIMO.Drawing;

/// <summary>
/// Header-level image metadata used by Office Open XML packages.
/// </summary>
public sealed class OfficeImageInfo {
    /// <summary>
    /// Creates a new image metadata value.
    /// </summary>
    public OfficeImageInfo(OfficeImageFormat format, int width, int height, double dpiX = 96.0, double dpiY = 96.0) {
        Format = format;
        Width = width;
        Height = height;
        DpiX = dpiX > 0 ? dpiX : 96.0;
        DpiY = dpiY > 0 ? dpiY : 96.0;
    }

    /// <summary>Image format.</summary>
    public OfficeImageFormat Format { get; }

    /// <summary>Width in pixels, or 0 when unknown.</summary>
    public int Width { get; }

    /// <summary>Height in pixels, or 0 when unknown.</summary>
    public int Height { get; }

    /// <summary>Horizontal resolution in DPI. Defaults to 96 when absent.</summary>
    public double DpiX { get; }

    /// <summary>Vertical resolution in DPI. Defaults to 96 when absent.</summary>
    public double DpiY { get; }

    /// <summary>Default MIME type for the detected format.</summary>
    public string MimeType => Format switch {
        OfficeImageFormat.Png => "image/png",
        OfficeImageFormat.Jpeg => "image/jpeg",
        OfficeImageFormat.Gif => "image/gif",
        OfficeImageFormat.Bmp => "image/bmp",
        OfficeImageFormat.Tiff => "image/tiff",
        OfficeImageFormat.Svg => "image/svg+xml",
        OfficeImageFormat.Emf => "image/x-emf",
        OfficeImageFormat.Wmf => "image/x-wmf",
        OfficeImageFormat.Icon => "image/x-icon",
        OfficeImageFormat.Pcx => "image/x-pcx",
        _ => "application/octet-stream"
    };
}
