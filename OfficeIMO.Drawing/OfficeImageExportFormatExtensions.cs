using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared metadata for dependency-free image export formats.
/// </summary>
public static class OfficeImageExportFormatExtensions {
    /// <summary>Returns the conventional file extension, including the leading dot.</summary>
    public static string GetFileExtension(this OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => ".png",
        OfficeImageExportFormat.Svg => ".svg",
        OfficeImageExportFormat.Jpeg => ".jpg",
        OfficeImageExportFormat.Tiff => ".tiff",
        OfficeImageExportFormat.Webp => ".webp",
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    /// <summary>Returns the Internet media type for the encoded image.</summary>
    public static string GetMimeType(this OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => "image/png",
        OfficeImageExportFormat.Svg => "image/svg+xml",
        OfficeImageExportFormat.Jpeg => "image/jpeg",
        OfficeImageExportFormat.Tiff => "image/tiff",
        OfficeImageExportFormat.Webp => "image/webp",
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    /// <summary>Returns whether a path extension is conventional for this output format.</summary>
    public static bool HasFileExtension(this OfficeImageExportFormat format, string? extension) {
        if (string.IsNullOrWhiteSpace(extension)) return false;
        string normalized = extension![0] == '.' ? extension : "." + extension;
        return format switch {
            OfficeImageExportFormat.Png => normalized.Equals(".png", StringComparison.OrdinalIgnoreCase),
            OfficeImageExportFormat.Svg => normalized.Equals(".svg", StringComparison.OrdinalIgnoreCase),
            OfficeImageExportFormat.Jpeg => normalized.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                            normalized.Equals(".jpeg", StringComparison.OrdinalIgnoreCase),
            OfficeImageExportFormat.Tiff => normalized.Equals(".tif", StringComparison.OrdinalIgnoreCase) ||
                                            normalized.Equals(".tiff", StringComparison.OrdinalIgnoreCase),
            OfficeImageExportFormat.Webp => normalized.Equals(".webp", StringComparison.OrdinalIgnoreCase),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
    }

    /// <summary>Returns whether the format is backed by raster pixels.</summary>
    public static bool IsRaster(this OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => true,
        OfficeImageExportFormat.Jpeg => true,
        OfficeImageExportFormat.Tiff => true,
        OfficeImageExportFormat.Webp => true,
        OfficeImageExportFormat.Svg => false,
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };
}
