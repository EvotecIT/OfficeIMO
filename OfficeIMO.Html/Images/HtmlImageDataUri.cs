using System.Text;

namespace OfficeIMO.Html;

/// <summary>
/// Represents an image data URI split into media type, encoding metadata, and payload.
/// </summary>
public sealed class HtmlImageDataUri {
    private HtmlImageDataUri(string metadata, string mediaType, string data, bool isBase64) {
        Metadata = metadata;
        MediaType = mediaType;
        Data = data;
        IsBase64 = isBase64;
    }

    /// <summary>
    /// Data URI metadata without the leading <c>data:</c> prefix.
    /// </summary>
    public string Metadata { get; }

    /// <summary>
    /// Declared image media type, for example <c>image/png</c>.
    /// </summary>
    public string MediaType { get; }

    /// <summary>
    /// Raw payload after the comma separator.
    /// </summary>
    public string Data { get; }

    /// <summary>
    /// Indicates whether the payload is base64 encoded.
    /// </summary>
    public bool IsBase64 { get; }

    /// <summary>
    /// Suggested file extension for the media type, including the leading dot.
    /// </summary>
    public string FileExtension => GetImageExtension(MediaType);

    /// <summary>
    /// Gets the suggested file extension for an image media type, including the leading dot.
    /// </summary>
    public static string GetFileExtension(string mediaType) => GetImageExtension(mediaType);

    /// <summary>
    /// Tries to parse an image data URI.
    /// </summary>
    public static bool TryParse(string? source, out HtmlImageDataUri dataUri) {
        dataUri = null!;
        if (string.IsNullOrWhiteSpace(source) || !source!.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        int commaIndex = source.IndexOf(',');
        if (commaIndex <= "data:".Length) {
            return false;
        }

        string metadata = source.Substring("data:".Length, commaIndex - "data:".Length);
        string mediaType = GetDataUriContentType(metadata);
        if (mediaType.Length == 0 || !mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        string data = source.Substring(commaIndex + 1);
        bool isBase64 = metadata.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0;
        dataUri = new HtmlImageDataUri(metadata, mediaType, data, isBase64);
        return true;
    }

    /// <summary>
    /// Decodes the image payload as bytes.
    /// </summary>
    public byte[] DecodeBytes() {
        if (!IsBase64) {
            return Encoding.UTF8.GetBytes(Uri.UnescapeDataString(Data));
        }

        string payload = Uri.UnescapeDataString(Data).Trim();
        return Convert.FromBase64String(payload);
    }

    /// <summary>
    /// Attempts to decode the image payload as bytes.
    /// </summary>
    public bool TryDecodeBytes(out byte[] bytes) {
        bytes = Array.Empty<byte>();
        try {
            bytes = DecodeBytes();
            return bytes.Length > 0;
        } catch (FormatException) {
            return false;
        }
    }

    /// <summary>
    /// Decodes the payload as UTF-8 text.
    /// </summary>
    public string DecodeText() {
        return IsBase64
            ? Encoding.UTF8.GetString(DecodeBytes())
            : Uri.UnescapeDataString(Data);
    }

    /// <summary>
    /// Estimates the decoded byte count without allocating decoded content when possible.
    /// </summary>
    public long EstimateDecodedByteCount() {
        if (!IsBase64) {
            return Encoding.UTF8.GetByteCount(Uri.UnescapeDataString(Data));
        }

        string payload = Uri.UnescapeDataString(Data).Trim();
        int length = payload.Length;
        int padding = 0;
        if (length > 0 && payload[length - 1] == '=') {
            padding++;
        }

        if (length > 1 && payload[length - 2] == '=') {
            padding++;
        }

        return (long)Math.Ceiling(length / 4D) * 3L - padding;
    }

    private static string GetDataUriContentType(string metadata) {
        if (string.IsNullOrWhiteSpace(metadata)) {
            return string.Empty;
        }

        int separatorIndex = metadata.IndexOf(';');
        string contentType = separatorIndex >= 0 ? metadata.Substring(0, separatorIndex) : metadata;
        return string.IsNullOrWhiteSpace(contentType) ? string.Empty : contentType.Trim();
    }

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
