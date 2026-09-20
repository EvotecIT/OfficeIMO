using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfEncoding {
    // Preserve the one-byte-to-one-character mapping on older target frameworks.
    public static string Latin1GetString(byte[] bytes) {
#if NET8_0_OR_GREATER
        return Encoding.Latin1.GetString(bytes);
#else
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = (char)bytes[i];
        return new string(chars);
#endif
    }

    public static string Latin1GetString(byte[] bytes, int index, int count) {
#if NET8_0_OR_GREATER
        return Encoding.Latin1.GetString(bytes, index, count);
#else
        var chars = new char[count];
        for (int i = 0; i < count; i++) chars[i] = (char)bytes[index + i];
        return new string(chars);
#endif
    }

    public static byte[] Latin1GetBytes(string s) {
        var bytes = new byte[s.Length];
        for (int i = 0; i < s.Length; i++) bytes[i] = (byte)(s[i] & 0xFF);
        return bytes;
    }

    // Encodes a StringBuilder's content to Latin1 bytes without materializing an intermediate string.
    // Page content streams are large and built in a StringBuilder, so skipping the ToString() saves a
    // full-length string allocation per page.
    public static byte[] Latin1GetBytes(StringBuilder sb) {
#if NET6_0_OR_GREATER
        var bytes = new byte[sb.Length];
        int pos = 0;
        foreach (System.ReadOnlyMemory<char> chunk in sb.GetChunks()) {
            System.ReadOnlySpan<char> span = chunk.Span;
            for (int i = 0; i < span.Length; i++) bytes[pos++] = (byte)(span[i] & 0xFF);
        }
        return bytes;
#else
        // GetChunks is unavailable here and the StringBuilder indexer walks the chunk linked list per
        // access (super-linear), so decode via a single string, matching the original ToString() path.
        return Latin1GetBytes(sb.ToString());
#endif
    }
}

