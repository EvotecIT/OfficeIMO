using System.Text;

namespace OfficeIMO.Html;

/// <summary>
/// Reads and decodes the CSS identifier subset used by generated-content contracts.
/// </summary>
internal static class HtmlCssIdentifierParser {
    internal static bool TryRead(string text, ref int cursor, out string value) {
        int start = cursor;
        if (!WouldStartIdentifier(text, cursor)) {
            value = string.Empty;
            return false;
        }
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

    private static bool WouldStartIdentifier(string text, int cursor) {
        if (cursor >= text.Length) return false;
        char first = text[cursor];
        if (first == '\\') return IsValidEscape(text, cursor);
        if (first != '-') return IsIdentifierStart(first);
        if (cursor + 1 >= text.Length) return false;
        char second = text[cursor + 1];
        return IsIdentifierStart(second) || second == '\\' && IsValidEscape(text, cursor + 1);
    }

    private static bool IsValidEscape(string text, int cursor) =>
        cursor + 1 < text.Length && text[cursor] == '\\' && text[cursor + 1] != '\n' && text[cursor + 1] != '\r' && text[cursor + 1] != '\f';
}
