namespace OfficeIMO.Pdf;

internal static class PdfStringParser {
    // Parses PDF literal string content (without surrounding parentheses) into original bytes (respecting escapes).
    public static byte[] ParseLiteralToBytes(string inner) =>
        ParseLiteralToBytes(inner, 0, inner.Length);

    internal static byte[] ParseLiteralToBytes(string source, int start, int length,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (length == 0) return Array.Empty<byte>();
        if (start < 0 || length < 0 || start > source.Length - length) throw new ArgumentOutOfRangeException(nameof(start));

        // Escapes can only remove characters, never add bytes. Most literal strings
        // contain no escapes, so this buffer is already the exact returned size.
        var bytes = new byte[length];
        int count = 0;
        int end = start + length;
        for (int i = start; i < end; i++) {
            if (((i - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char c = source[i];
            if (c == '\\') {
                if (i + 1 >= end) break;
                char n = source[++i];
                switch (n) {
                    case 'n': bytes[count++] = (byte)'\n'; break;
                    case 'r': bytes[count++] = (byte)'\r'; break;
                    case 't': bytes[count++] = (byte)'\t'; break;
                    case 'b': bytes[count++] = (byte)'\b'; break;
                    case 'f': bytes[count++] = (byte)'\f'; break;
                    case '\\': bytes[count++] = (byte)'\\'; break;
                    case '(': bytes[count++] = (byte)'('; break;
                    case ')': bytes[count++] = (byte)')'; break;
                    case '\n': /* line continuation */ break;
                    case '\r':
                        if (i + 1 < end && source[i + 1] == '\n') {
                            i++;
                        }
                        break;
                    default:
                        if (IsOctalDigit(n)) {
                            int v = n - '0';
                            // up to 2 more octal digits
                            for (int k = 0; k < 2 && i + 1 < end && IsOctalDigit(source[i + 1]); k++) {
                                v = (v << 3) + (source[++i] - '0');
                            }
                            bytes[count++] = (byte)(v & 0xFF);
                        } else {
                            bytes[count++] = (byte)(n & 0xFF);
                        }
                        break;
                }
            } else {
                bytes[count++] = (byte)(c & 0xFF);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (count == bytes.Length) return bytes;
        var result = new byte[count];
        for (int offset = 0; offset < count;) {
            cancellationToken.ThrowIfCancellationRequested();
            int chunkLength = Math.Min(64 * 1024, count - offset);
            Buffer.BlockCopy(bytes, offset, result, offset, chunkLength);
            offset += chunkLength;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private static bool IsOctalDigit(char c) => c >= '0' && c <= '7';
}

