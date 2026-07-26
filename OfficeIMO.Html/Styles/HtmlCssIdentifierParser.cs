using System.Text;

namespace OfficeIMO.Html;

/// <summary>
/// Reads and decodes the CSS identifier subset used by generated-content contracts.
/// </summary>
internal static class HtmlCssIdentifierParser {
    internal static bool TryRead(string text, ref int cursor, out string value) {
        int start = cursor;
        var result = new StringBuilder();
        bool first = true;
        while (cursor < text.Length) {
            char current = text[cursor];
            if (current == '\\') {
                if (!HtmlCssEscapeDecoder.TryDecodeEscape(
                        text,
                        cursor,
                        out string decoded,
                        out int consumedCharacters)
                    || decoded.Length == 0) {
                    cursor = start;
                    value = string.Empty;
                    return false;
                }
                result.Append(decoded);
                cursor += consumedCharacters;
                first = false;
                continue;
            }
            if (first ? !IsIdentifierStart(current) : !IsIdentifierCharacter(current)) break;
            result.Append(current);
            cursor++;
            first = false;
        }

        if (first) {
            cursor = start;
            value = string.Empty;
            return false;
        }
        value = result.ToString();
        return true;
    }

    internal static bool TryParse(string text, out string value) {
        int cursor = 0;
        if (!TryRead(text, ref cursor, out value) || cursor != text.Length) {
            value = string.Empty;
            return false;
        }
        return true;
    }

    private static bool IsIdentifierStart(char value) =>
        char.IsLetter(value) || value == '_' || value == '-' || value >= 0x80;

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '_' || value == '-' || value >= 0x80;
}
