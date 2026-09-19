namespace OfficeIMO.Pdf;

internal static class PdfStringParser {
    // Parses PDF literal string content (without surrounding parentheses) into original bytes (respecting escapes).
    public static byte[] ParseLiteralToBytes(string inner) {
        if (inner.Length == 0) return Array.Empty<byte>();

        // Escapes can only remove characters, never add bytes. Most literal strings
        // contain no escapes, so this buffer is already the exact returned size.
        var bytes = new byte[inner.Length];
        int count = 0;
        for (int i = 0; i < inner.Length; i++) {
            char c = inner[i];
            if (c == '\\') {
                if (i + 1 >= inner.Length) break;
                char n = inner[++i];
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
                        if (i + 1 < inner.Length && inner[i + 1] == '\n') {
                            i++;
                        }
                        break;
                    default:
                        if (IsOctalDigit(n)) {
                            int v = n - '0';
                            // up to 2 more octal digits
                            for (int k = 0; k < 2 && i + 1 < inner.Length && IsOctalDigit(inner[i + 1]); k++) {
                                v = (v << 3) + (inner[++i] - '0');
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

        if (count == bytes.Length) return bytes;
        Array.Resize(ref bytes, count);
        return bytes;
    }

    private static bool IsOctalDigit(char c) => c >= '0' && c <= '7';
}

