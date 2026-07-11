namespace OfficeIMO.Html;

/// <summary>
/// Represents an image data URI split into media type, encoding metadata, and payload.
/// </summary>
public sealed class HtmlImageDataUri {
    private readonly HtmlDataUri _dataUri;

    private HtmlImageDataUri(HtmlDataUri dataUri) {
        _dataUri = dataUri;
    }

    /// <summary>Data URI metadata without the leading <c>data:</c> prefix.</summary>
    public string Metadata => _dataUri.Metadata;

    /// <summary>Declared image media type, for example <c>image/png</c>.</summary>
    public string MediaType => _dataUri.MediaType;

    /// <summary>Raw payload after the comma separator.</summary>
    public string Data => _dataUri.Data;

    /// <summary>Indicates whether the payload is base64 encoded.</summary>
    public bool IsBase64 => _dataUri.IsBase64;

    /// <summary>Suggested file extension for the media type, including the leading dot.</summary>
    public string FileExtension => GetImageExtension(MediaType);

    /// <summary>Gets the suggested file extension for an image media type.</summary>
    public static string GetFileExtension(string mediaType) => GetImageExtension(mediaType);

    /// <summary>Tries to parse an image data URI.</summary>
    public static bool TryParse(string? source, out HtmlImageDataUri dataUri) {
        dataUri = null!;
        if (!HtmlDataUri.TryParse(source, out HtmlDataUri parsed)
            || !parsed.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        dataUri = new HtmlImageDataUri(parsed);
        return true;
    }

    /// <summary>Decodes the image payload as bytes.</summary>
    public byte[] DecodeBytes() => _dataUri.DecodeBytes();

    /// <summary>Attempts to decode the image payload as bytes.</summary>
    public bool TryDecodeBytes(out byte[] bytes) => _dataUri.TryDecodeBytes(out bytes);

    /// <summary>Decodes the payload as UTF-8 text.</summary>
    public string DecodeText() => _dataUri.DecodeText();

    /// <summary>Estimates decoded byte count without allocating decoded content when possible.</summary>
    public long EstimateDecodedByteCount() => _dataUri.EstimateDecodedByteCount();

    private static string GetImageExtension(string mediaType) {
        return mediaType.ToLowerInvariant() switch {
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/png" => ".png",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/webp" => ".webp",
            "image/svg+xml" => ".svg",
            _ => ".bin"
        };
    }
}
