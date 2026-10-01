namespace OfficeIMO.Rtf;

internal static class RtfListTextCodec {
    internal static string DecodeText(string text) {
        string content = TrimTerminator(text);
        if (content.Length > 0) content = content.Substring(1);
        var result = new StringBuilder();
        foreach (char character in content) {
            if (character <= 8) result.Append('%').Append((character + 1).ToString(CultureInfo.InvariantCulture));
            else result.Append(character);
        }
        return result.ToString();
    }

    internal static string DecodeNumbers(string text) => TrimTerminator(text);

    internal static string EncodeText(string text, out string offsets) {
        var encoded = new StringBuilder();
        var positions = new StringBuilder();
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '%' && i + 1 < text.Length && text[i + 1] is >= '1' and <= '9') {
                positions.Append((char)(encoded.Length + 1));
                encoded.Append((char)(text[++i] - '1'));
            } else encoded.Append(text[i]);
        }
        offsets = positions.ToString();
        return encoded.ToString();
    }

    private static string TrimTerminator(string text) => text.EndsWith(";", StringComparison.Ordinal) ? text.Substring(0, text.Length - 1) : text;
}
