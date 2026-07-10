using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal static class HtmlCssEscapeDecoder {
    internal static string Decode(string source) {
        if (string.IsNullOrEmpty(source) || source.IndexOf('\\') < 0) return source;

        var result = new StringBuilder(source.Length);
        for (int index = 0; index < source.Length; index++) {
            char current = source[index];
            if (current != '\\' || index + 1 >= source.Length) {
                result.Append(current);
                continue;
            }

            int cursor = index + 1;
            if (source[cursor] == '\r' || source[cursor] == '\n' || source[cursor] == '\f') {
                if (source[cursor] == '\r' && cursor + 1 < source.Length && source[cursor + 1] == '\n') cursor++;
                index = cursor;
                continue;
            }

            int hexStart = cursor;
            while (cursor < source.Length && cursor - hexStart < 6 && IsHexDigit(source[cursor])) cursor++;
            if (cursor > hexStart) {
                string hex = source.Substring(hexStart, cursor - hexStart);
                if (!int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int codePoint)
                    || codePoint == 0
                    || codePoint > 0x10FFFF
                    || codePoint >= 0xD800 && codePoint <= 0xDFFF) {
                    result.Append('\uFFFD');
                } else {
                    result.Append(char.ConvertFromUtf32(codePoint));
                }

                if (cursor < source.Length && char.IsWhiteSpace(source[cursor])) cursor++;
                index = cursor - 1;
                continue;
            }

            result.Append(source[cursor]);
            index = cursor;
        }

        return result.ToString();
    }

    private static bool IsHexDigit(char value) =>
        value >= '0' && value <= '9'
        || value >= 'a' && value <= 'f'
        || value >= 'A' && value <= 'F';
}
