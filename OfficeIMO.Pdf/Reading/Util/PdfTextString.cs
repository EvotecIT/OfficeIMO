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

    internal static byte[] DecodeHexBytes(string raw) {
        var hex = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++) {
            char ch = raw[i];
            if (!char.IsWhiteSpace(ch)) {
                hex.Append(ch);
            }
        }

        if ((hex.Length & 1) == 1) {
            hex.Append('0');
        }

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }

        return bytes;
    }

    private static int HexNibble(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
        if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
        throw new FormatException($"Invalid hex character '{c}'.");
    }
}
