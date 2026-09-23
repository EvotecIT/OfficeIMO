using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfEncoding {
    internal static string DecodeCancellable(Encoding encoding, byte[] bytes, int index, int count,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return encoding.GetString(bytes, index, count);

        var decoder = encoding.GetDecoder();
        var chars = new char[8192];
        var builder = new StringBuilder(count);
        int end = index + count;
        while (index < end) {
            cancellationToken.ThrowIfCancellationRequested();
            int chunk = Math.Min(4096, end - index);
            int charsUsed = decoder.GetChars(bytes, index, chunk, chars, 0, index + chunk == end);
            builder.Append(chars, 0, charsUsed);
            index += chunk;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }

    internal static string Latin1GetStringCancellable(byte[] bytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return Latin1GetString(bytes);
#if NET8_0_OR_GREATER
        return string.Create(bytes.Length, (Bytes: bytes, Token: cancellationToken), (chars, state) => {
            for (int i = 0; i < state.Bytes.Length; i++) {
                if ((i & 4095) == 0) state.Token.ThrowIfCancellationRequested();
                chars[i] = (char)state.Bytes[i];
            }
            state.Token.ThrowIfCancellationRequested();
        });
#else
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            chars[i] = (char)bytes[i];
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new string(chars);
#endif
    }

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

    internal static string Latin1GetStringCancellable(byte[] bytes, int index, int count,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return Latin1GetString(bytes, index, count);
        var chars = new char[count];
        for (int i = 0; i < count; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            chars[i] = (char)bytes[index + i];
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new string(chars);
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

