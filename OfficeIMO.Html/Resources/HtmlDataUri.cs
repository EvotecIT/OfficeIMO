using System.IO;
using System.Text;

namespace OfficeIMO.Html;

/// <summary>Media-neutral data URI with bounded-size inspection and byte decoding.</summary>
public sealed class HtmlDataUri {
    private HtmlDataUri(string metadata, string mediaType, string data, bool isBase64) {
        Metadata = metadata;
        MediaType = mediaType;
        Data = data;
        IsBase64 = isBase64;
    }

    /// <summary>Metadata without the leading <c>data:</c> prefix.</summary>
    public string Metadata { get; }

    /// <summary>Declared media type.</summary>
    public string MediaType { get; }

    /// <summary>Raw payload after the comma separator.</summary>
    public string Data { get; }

    /// <summary>Whether the payload uses base64 encoding.</summary>
    public bool IsBase64 { get; }

    /// <summary>Tries to parse a data URI with an explicit media type.</summary>
    public static bool TryParse(string? source, out HtmlDataUri dataUri) {
        dataUri = null!;
        if (string.IsNullOrWhiteSpace(source) || !source!.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        int commaIndex = source.IndexOf(',');
        if (commaIndex <= "data:".Length) {
            return false;
        }

        string metadata = source.Substring("data:".Length, commaIndex - "data:".Length);
        string mediaType = GetContentType(metadata);
        if (mediaType.Length == 0) {
            return false;
        }

        dataUri = new HtmlDataUri(
            metadata,
            mediaType,
            source.Substring(commaIndex + 1),
            HasBase64Flag(metadata));
        return true;
    }

    /// <summary>Decodes the payload as bytes.</summary>
    public byte[] DecodeBytes() {
        if (!IsBase64) {
            return DecodePercentEncodedBytes(Data);
        }

        return Convert.FromBase64String(NormalizeBase64Payload(Uri.UnescapeDataString(Data)));
    }

    /// <summary>Attempts to decode the payload as bytes.</summary>
    public bool TryDecodeBytes(out byte[] bytes) {
        bytes = Array.Empty<byte>();
        try {
            bytes = DecodeBytes();
            return bytes.Length > 0;
        } catch (FormatException) {
            return false;
        }
    }

    /// <summary>Decodes the payload as UTF-8 text.</summary>
    public string DecodeText() => IsBase64
        ? Encoding.UTF8.GetString(DecodeBytes())
        : Uri.UnescapeDataString(Data);

    /// <summary>Calculates decoded byte count without allocating the decoded payload.</summary>
    public long EstimateDecodedByteCount() {
        if (!IsBase64) {
            return CountPercentDecodedBytes(Data);
        }

        string payload = NormalizeBase64Payload(Uri.UnescapeDataString(Data));
        if (payload.Length == 0 || payload.Length % 4 != 0) {
            throw new FormatException("Invalid base64 data URI payload length.");
        }

        int padding = payload.EndsWith("==", StringComparison.Ordinal) ? 2
            : payload.EndsWith("=", StringComparison.Ordinal) ? 1
            : 0;
        return ((long)payload.Length / 4L * 3L) - padding;
    }

    private static byte[] DecodePercentEncodedBytes(string data) {
        using var stream = new MemoryStream();
        var text = new StringBuilder();
        for (int index = 0; index < data.Length; index++) {
            char character = data[index];
            if (character != '%') {
                text.Append(character);
                continue;
            }

            FlushTextBytes(text, stream);
            stream.WriteByte(ReadEscapedByte(data, index));
            index += 2;
        }

        FlushTextBytes(text, stream);
        return stream.ToArray();
    }

    private static long CountPercentDecodedBytes(string data) {
        long count = 0L;
        var text = new StringBuilder();
        for (int index = 0; index < data.Length; index++) {
            char character = data[index];
            if (character != '%') {
                text.Append(character);
                continue;
            }

            count += Encoding.UTF8.GetByteCount(text.ToString());
            text.Clear();
            _ = ReadEscapedByte(data, index);
            count++;
            index += 2;
        }

        count += Encoding.UTF8.GetByteCount(text.ToString());
        return count;
    }

    private static void FlushTextBytes(StringBuilder text, Stream stream) {
        if (text.Length == 0) {
            return;
        }

        byte[] bytes = Encoding.UTF8.GetBytes(text.ToString());
        stream.Write(bytes, 0, bytes.Length);
        text.Clear();
    }

    private static byte ReadEscapedByte(string data, int percentIndex) {
        if (percentIndex + 2 >= data.Length
            || !TryReadHex(data[percentIndex + 1], out byte high)
            || !TryReadHex(data[percentIndex + 2], out byte low)) {
            throw new UriFormatException("Invalid percent escape in data URI payload.");
        }

        return (byte)((high << 4) | low);
    }

    private static bool TryReadHex(char value, out byte nibble) {
        if (value >= '0' && value <= '9') {
            nibble = (byte)(value - '0');
            return true;
        }

        if (value >= 'A' && value <= 'F') {
            nibble = (byte)(value - 'A' + 10);
            return true;
        }

        if (value >= 'a' && value <= 'f') {
            nibble = (byte)(value - 'a' + 10);
            return true;
        }

        nibble = 0;
        return false;
    }

    private static string GetContentType(string metadata) {
        int separatorIndex = metadata.IndexOf(';');
        string contentType = separatorIndex >= 0 ? metadata.Substring(0, separatorIndex) : metadata;
        return string.IsNullOrWhiteSpace(contentType) ? string.Empty : contentType.Trim();
    }

    private static bool HasBase64Flag(string metadata) =>
        metadata.Split(';').Any(part => part.Trim().Equals("base64", StringComparison.OrdinalIgnoreCase));

    private static string NormalizeBase64Payload(string payload) {
        var builder = new StringBuilder(payload.Length);
        foreach (char character in payload.Trim()) {
            if (!char.IsWhiteSpace(character)) {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }
}
