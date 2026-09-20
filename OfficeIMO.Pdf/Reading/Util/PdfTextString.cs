using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfTextString {
    internal static int GetDecodedCharacterCount(byte[] bytes) {
        if (bytes == null || bytes.Length == 0) {
            return 0;
        }

        if (bytes.Length >= 2 &&
            (bytes[0] == 0xFE && bytes[1] == 0xFF || bytes[0] == 0xFF && bytes[1] == 0xFE)) {
            return (bytes.Length - 2) / 2;
        }

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
            return Encoding.UTF8.GetCharCount(bytes, 3, bytes.Length - 3);
        }

        return bytes.Length;
    }

    public static string Decode(byte[] bytes) {
        if (bytes == null || bytes.Length == 0) {
            return string.Empty;
        }

        if (bytes.Length >= 2) {
            if (bytes[0] == 0xFE && bytes[1] == 0xFF) {
                return DecodeUtf16BigEndian(bytes, 2);
            }

            if (bytes[0] == 0xFF && bytes[1] == 0xFE) {
                return DecodeUtf16LittleEndian(bytes, 2);
            }
        }

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        return PdfWinAnsiEncoding.Decode(bytes);
    }

    public static byte[] Encode(string value) {
        if (string.IsNullOrEmpty(value)) {
            return Array.Empty<byte>();
        }

        if (PdfWinAnsiEncoding.CanEncode(value, out _)) {
            return PdfWinAnsiEncoding.Encode(value);
        }

        var result = new byte[2 + (value.Length * 2)];
        result[0] = 0xFE;
        result[1] = 0xFF;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            result[2 + (i * 2)] = (byte)(ch >> 8);
            result[3 + (i * 2)] = (byte)(ch & 0xFF);
        }

        return result;
    }

    public static string DecodeHex(string raw) {
        if (string.IsNullOrWhiteSpace(raw)) {
            return string.Empty;
        }

        return Decode(DecodeHexBytes(raw));
    }

    public static string DecodeLiteral(string inner) {
        if (string.IsNullOrEmpty(inner)) {
            return string.Empty;
        }

        return Decode(PdfStringParser.ParseLiteralToBytes(inner));
    }

    private static string DecodeUtf16BigEndian(byte[] bytes, int offset) {
        var builder = new StringBuilder((bytes.Length - offset) / 2);
        for (int i = offset; i + 1 < bytes.Length; i += 2) {
            builder.Append((char)((bytes[i] << 8) | bytes[i + 1]));
        }

        return builder.ToString();
    }

    private static string DecodeUtf16LittleEndian(byte[] bytes, int offset) {
        var builder = new StringBuilder((bytes.Length - offset) / 2);
        for (int i = offset; i + 1 < bytes.Length; i += 2) {
            builder.Append((char)(bytes[i] | (bytes[i + 1] << 8)));
        }

        return builder.ToString();
    }

    internal static byte[] DecodeHexBytes(string raw) => DecodeHexBytes(raw, 0, raw.Length);

    internal static byte[] DecodeHexBytes(string source, int start, int length) {
        if (start < 0 || length < 0 || start > source.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        int end = start + length;
        int nibbleCount = 0;
        for (int i = start; i < end; i++) {
            if (!char.IsWhiteSpace(source[i])) nibbleCount++;
        }

        if (nibbleCount == 0) return Array.Empty<byte>();
        var bytes = new byte[(nibbleCount + 1) / 2];
        int highNibble = -1;
        int byteIndex = 0;
        for (int i = start; i < end; i++) {
            char ch = source[i];
            if (char.IsWhiteSpace(ch)) continue;
            int nibble = HexNibble(ch);
            if (highNibble < 0) {
                highNibble = nibble;
            } else {
                bytes[byteIndex++] = (byte)((highNibble << 4) | nibble);
                highNibble = -1;
            }
        }

        if (highNibble >= 0) bytes[byteIndex] = (byte)(highNibble << 4);
        return bytes;
    }

    private static int HexNibble(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
        if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
        throw new FormatException($"Invalid hex character '{c}'.");
    }
}
