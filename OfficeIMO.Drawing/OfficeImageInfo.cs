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
    public string MimeType => GetMimeType(Format);

    /// <summary>
    /// Returns the default MIME type for a known image format.
    /// </summary>
    /// <param name="format">Image format.</param>
    /// <returns>The default MIME content type, or application/octet-stream for unknown formats.</returns>
    public static string GetMimeType(OfficeImageFormat format) => format switch {
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
        OfficeImageFormat.Webp => "image/webp",
        _ => "application/octet-stream"
    };

    /// <summary>
    /// Maps a MIME content type to a known image format.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <returns>The matching image format, or <see cref="OfficeImageFormat.Unknown" /> when unsupported.</returns>
    public static OfficeImageFormat FromMimeType(string? contentType) {
        if (string.IsNullOrWhiteSpace(contentType)) {
            return OfficeImageFormat.Unknown;
        }

        int separator = contentType!.IndexOf(';');
        string normalized = separator >= 0
            ? contentType.Substring(0, separator)
            : contentType;
        normalized = normalized.Trim().ToLowerInvariant();
        return normalized switch {
            "image/png" => OfficeImageFormat.Png,
            "image/jpeg" or "image/jpg" or "image/pjpeg" => OfficeImageFormat.Jpeg,
            "image/gif" => OfficeImageFormat.Gif,
            "image/bmp" or "image/x-bmp" => OfficeImageFormat.Bmp,
            "image/tiff" or "image/tif" => OfficeImageFormat.Tiff,
            "image/svg+xml" or "image/svg" => OfficeImageFormat.Svg,
            "image/x-emf" or "image/emf" => OfficeImageFormat.Emf,
            "image/x-wmf" or "image/wmf" => OfficeImageFormat.Wmf,
            "image/x-icon" or "image/vnd.microsoft.icon" => OfficeImageFormat.Icon,
            "image/x-pcx" => OfficeImageFormat.Pcx,
            "image/webp" => OfficeImageFormat.Webp,
            _ => OfficeImageFormat.Unknown
        };
    }
}
