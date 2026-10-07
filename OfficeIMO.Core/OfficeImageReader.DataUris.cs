using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    /// <summary>Decodes and validates a bounded base64 PNG or JPEG data URI.</summary>
    /// <remarks>Rejects unsupported MIME types, malformed or oversized payloads, and content that does not match the declared MIME type.</remarks>
    public static bool TryReadBase64DataUri(string? source, long maximumImageBytes,
        out byte[] bytes, out OfficeImageInfo info) {
        bytes = Array.Empty<byte>();
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (source == null || maximumImageBytes < 0
            || !source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return false;
        int comma = source.IndexOf(',');
        if (comma < 0) return false;
        string metadata = source.Substring(5, comma - 5);
        string mime;
        if (string.Equals(metadata, "image/png;base64", StringComparison.OrdinalIgnoreCase)) mime = "image/png";
        else if (string.Equals(metadata, "image/jpeg;base64", StringComparison.OrdinalIgnoreCase)) mime = "image/jpeg";
        else return false;
        long payloadLength = source.Length - comma - 1L;
        // Four encoded characters represent at most three bytes, with up to two padding bytes.
        if ((payloadLength + 3L) / 4L * 3L - 2L > maximumImageBytes) return false;
        try {
            bytes = Convert.FromBase64String(source.Substring(comma + 1));
        } catch (FormatException) {
            return false;
        }
        if (bytes.LongLength > maximumImageBytes
            || !TryValidateContent(bytes, null, out info)
            || !string.Equals(info.MimeType, mime, StringComparison.OrdinalIgnoreCase)) {
            bytes = Array.Empty<byte>();
            return false;
        }
        return true;
    }
}
