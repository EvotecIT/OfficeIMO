using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Canonical OpenDocument image extension and media-type mappings.</summary>
internal static class OdfImageFormats {
    internal static bool TryGetExtension(OfficeImageFormat format, out string extension) {
        extension = format switch {
            OfficeImageFormat.Png => ".png",
            OfficeImageFormat.Jpeg => ".jpg",
            OfficeImageFormat.Gif => ".gif",
            OfficeImageFormat.Svg => ".svg",
            OfficeImageFormat.Bmp => ".bmp",
            OfficeImageFormat.Webp => ".webp",
            _ => string.Empty
        };
        return extension.Length > 0;
    }

    internal static bool TryGetMediaType(string? pathOrExtension, out string mediaType) {
        string extension = NormalizeExtension(pathOrExtension);
        mediaType = extension switch {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".svg" => "image/svg+xml",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".webp" => "image/webp",
            _ => string.Empty
        };
        return mediaType.Length > 0;
    }

    internal static bool TryNormalizeStoredExtension(string? fileName, out string extension) {
        extension = NormalizeExtension(fileName);
        if (extension == ".jpeg") extension = ".jpg";
        return extension is ".png" or ".jpg" or ".gif" or ".svg" or ".bmp" or ".webp";
    }

    internal static bool TryGetFormat(string? mediaType, out OfficeImageFormat format) {
        format = OfficeImageInfo.FromMimeType(mediaType);
        return TryGetExtension(format, out _);
    }

    private static string NormalizeExtension(string? pathOrExtension) {
        if (string.IsNullOrWhiteSpace(pathOrExtension)) return string.Empty;
        string value = pathOrExtension!.Trim();
        string extension = value.StartsWith(".", StringComparison.Ordinal)
            ? value
            : Path.GetExtension(value);
        return extension.ToLowerInvariant();
    }
}
