using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal static class HtmlCssEscapeDecoder {
    internal static string Decode(string source) {
        if (string.IsNullOrEmpty(source) || source.IndexOf('\\') < 0) return source;

        var result = new StringBuilder(source.Length);
        for (int index = 0; index < source.Length; index++) {
            char current = source[index];
            if (current != '\\') {
                result.Append(current);
                continue;
            }

            TryDecodeEscape(source, index, out string decoded, out int consumedCharacters);
            result.Append(decoded);
            index += consumedCharacters - 1;
        }

        return result.ToString();
    }

    internal static bool TryDecodeEscape(
        string source,
        int backslashIndex,
        out string value,
        out int consumedCharacters) {
        value = string.Empty;
        consumedCharacters = 0;
        if (backslashIndex < 0
            || backslashIndex >= source.Length
            || source[backslashIndex] != '\\') {
            return false;
        }

        int cursor = backslashIndex + 1;
        if (cursor >= source.Length) {
            value = "\\";
            consumedCharacters = 1;
            return true;
        }

        if (source[cursor] == '\r' || source[cursor] == '\n' || source[cursor] == '\f') {
            cursor++;
            if (source[cursor - 1] == '\r' && cursor < source.Length && source[cursor] == '\n') cursor++;
            consumedCharacters = cursor - backslashIndex;
            return true;
        }

        int hexStart = cursor;
        while (cursor < source.Length && cursor - hexStart < 6 && IsHexDigit(source[cursor])) cursor++;
        if (cursor > hexStart) {
            string hex = source.Substring(hexStart, cursor - hexStart);
            if (!int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int codePoint)
                || codePoint == 0
                || codePoint > 0x10FFFF
                || codePoint >= 0xD800 && codePoint <= 0xDFFF) {
                value = "\uFFFD";
            } else {
                value = char.ConvertFromUtf32(codePoint);
            }

            if (cursor < source.Length && char.IsWhiteSpace(source[cursor])) {
                char terminator = source[cursor++];
                if (terminator == '\r' && cursor < source.Length && source[cursor] == '\n') cursor++;
            }
            consumedCharacters = cursor - backslashIndex;
            return true;
        }

        value = source[cursor].ToString();
        consumedCharacters = cursor - backslashIndex + 1;
        return true;
    }

    private static bool IsHexDigit(char value) =>
        value >= '0' && value <= '9'
        || value >= 'a' && value <= 'f'
        || value >= 'A' && value <= 'F';
}
