using System.Text;
using System.IO;

namespace OfficeIMO.Pdf.Filters;

internal static class AsciiHexDecoder {
    public static byte[] Decode(byte[] data) {
        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }

        var hex = new StringBuilder(data.Length);
        for (int i = 0; i < data.Length; i++) {
            char ch = (char)data[i];
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (ch == '>') {
                break;
            }

            hex.Append(ch);
        }

        if (hex.Length == 0) {
            return Array.Empty<byte>();
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

    public static bool TryDecode(byte[] data, int maxOutputBytes, out byte[] output) {
        output = Array.Empty<byte>();
        if (maxOutputBytes < 0) {
            return false;
        }

        if (data == null || data.Length == 0) {
            return true;
        }

        using var stream = new MemoryStream();
        int? highNibble = null;
        for (int i = 0; i < data.Length; i++) {
            char ch = (char)data[i];
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (ch == '>') {
                break;
            }

            int nibble = HexNibble(ch);
            if (highNibble is null) {
                highNibble = nibble;
                continue;
            }

            if (stream.Length >= maxOutputBytes) {
                return false;
            }

            stream.WriteByte((byte)((highNibble.Value << 4) | nibble));
            highNibble = null;
        }

        if (highNibble is not null) {
            if (stream.Length >= maxOutputBytes) {
                return false;
            }

            stream.WriteByte((byte)(highNibble.Value << 4));
        }

        output = stream.ToArray();
        return true;
    }

    private static int HexNibble(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
        if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
        throw new FormatException($"Invalid ASCIIHex character '{c}'.");
    }
}
