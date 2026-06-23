using System;

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
    /// Returns the canonical file extension for a known image format.
    /// </summary>
    /// <param name="format">Image format.</param>
    /// <returns>The default file extension including the leading dot, or .bin for unknown formats.</returns>
    public static string GetDefaultExtension(OfficeImageFormat format) => format switch {
        OfficeImageFormat.Png => ".png",
        OfficeImageFormat.Jpeg => ".jpeg",
        OfficeImageFormat.Gif => ".gif",
        OfficeImageFormat.Bmp => ".bmp",
        OfficeImageFormat.Tiff => ".tiff",
        OfficeImageFormat.Svg => ".svg",
        OfficeImageFormat.Emf => ".emf",
        OfficeImageFormat.Wmf => ".wmf",
        OfficeImageFormat.Icon => ".ico",
        OfficeImageFormat.Pcx => ".pcx",
        OfficeImageFormat.Webp => ".webp",
        _ => ".bin"
    };

    /// <summary>
    /// Returns the default MIME type for an image file name or extension.
    /// </summary>
    /// <param name="fileName">File name, path, or bare extension.</param>
    /// <returns>The default MIME content type, or application/octet-stream for unknown formats.</returns>
    public static string GetMimeTypeFromExtension(string? fileName) =>
        GetMimeType(OfficeImageReader.FromExtension(fileName));

    /// <summary>
    /// Returns whether the image format is safe for inline HTML preview galleries.
    /// </summary>
    /// <param name="format">Image format.</param>
    /// <returns><c>true</c> when the format can be previewed inline without embedding active SVG markup.</returns>
    public static bool IsBrowserPreviewSafeFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png ||
        format == OfficeImageFormat.Jpeg ||
        format == OfficeImageFormat.Gif ||
        format == OfficeImageFormat.Bmp ||
        format == OfficeImageFormat.Webp;

    /// <summary>
    /// Returns whether the image MIME content type is safe for inline HTML preview galleries.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <returns><c>true</c> when the content type maps to an inline-preview-safe image format.</returns>
    public static bool IsBrowserPreviewSafeContentType(string? contentType) =>
        IsBrowserPreviewSafeFormat(FromMimeType(contentType));

    /// <summary>
    /// Returns whether the image file name or extension is safe for inline HTML preview galleries.
    /// </summary>
    /// <param name="fileName">File name, path, or bare extension.</param>
    /// <returns><c>true</c> when the extension maps to an inline-preview-safe image format.</returns>
    public static bool IsBrowserPreviewSafeExtension(string? fileName) =>
        IsBrowserPreviewSafeFormat(OfficeImageReader.FromExtension(fileName));

    /// <summary>
    /// Maps a MIME content type to a known image format.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <returns>The matching image format, or <see cref="OfficeImageFormat.Unknown" /> when unsupported.</returns>
    public static OfficeImageFormat FromMimeType(string? contentType) {
        string normalized = NormalizeMimeType(contentType);
        if (string.IsNullOrEmpty(normalized)) {
            return OfficeImageFormat.Unknown;
        }

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

    /// <summary>
    /// Tries to normalize an image MIME content type by trimming parameters and canonicalizing known aliases.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <param name="normalizedContentType">Canonical image MIME type for known formats, or the normalized image MIME type for unknown image formats.</param>
    /// <returns><see langword="true" /> when the value is an image MIME type.</returns>
    public static bool TryNormalizeImageContentType(string? contentType, out string normalizedContentType) {
        string normalized = NormalizeMimeType(contentType);
        if (!normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
            normalizedContentType = string.Empty;
            return false;
        }

        OfficeImageFormat format = FromMimeType(normalized);
        normalizedContentType = format == OfficeImageFormat.Unknown
            ? normalized
            : GetMimeType(format);
        return true;
    }

    /// <summary>
    /// Normalizes a MIME content type by removing parameters and applying case-insensitive comparison casing.
    /// </summary>
    /// <param name="contentType">MIME content type, optionally with parameters.</param>
    /// <returns>The normalized MIME content type, or an empty string for missing values.</returns>
    public static string NormalizeMimeType(string? contentType) {
        if (string.IsNullOrWhiteSpace(contentType)) {
            return string.Empty;
        }

        int separator = contentType!.IndexOf(';');
        string normalized = separator >= 0
            ? contentType.Substring(0, separator)
            : contentType;
        return normalized.Trim().ToLowerInvariant();
    }
}
