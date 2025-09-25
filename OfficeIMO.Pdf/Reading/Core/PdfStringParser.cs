namespace OfficeIMO.Pdf;

internal static class PdfStringParser {
    // Parses PDF literal string content (without surrounding parentheses) into original bytes (respecting escapes).
    public static byte[] ParseLiteralToBytes(string inner) {
        var bytes = new List<byte>(inner.Length);
        for (int i = 0; i < inner.Length; i++) {
            char c = inner[i];
            if (c == '\\') {
                if (i + 1 >= inner.Length) break;
                char n = inner[++i];
                switch (n) {
                    case 'n': bytes.Add((byte)'\n'); break;
                    case 'r': bytes.Add((byte)'\r'); break;
                    case 't': bytes.Add((byte)'\t'); break;
                    case 'b': bytes.Add((byte)'\b'); break;
                    case 'f': bytes.Add((byte)'\f'); break;
                    case '\\': bytes.Add((byte)'\\'); break;
                    case '(': bytes.Add((byte)'('); break;
                    case ')': bytes.Add((byte)')'); break;
                    case '\n': /* line continuation */ break;
                    default:
                        if (IsOctalDigit(n)) {
                            int v = n - '0';
                            // up to 2 more octal digits
                            for (int k = 0; k < 2 && i + 1 < inner.Length && IsOctalDigit(inner[i + 1]); k++) {
                                v = (v << 3) + (inner[++i] - '0');
                            }
                            bytes.Add((byte)(v & 0xFF));
                        } else {
                            bytes.Add((byte)(n & 0xFF));
                        }
                        break;
                }
            } else {
                bytes.Add((byte)(c & 0xFF));
            }
        }
        return bytes.ToArray();
    }

    private static bool IsOctalDigit(char c) => c >= '0' && c <= '7';
}

