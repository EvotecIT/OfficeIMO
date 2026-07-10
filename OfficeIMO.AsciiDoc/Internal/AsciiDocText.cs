namespace OfficeIMO.AsciiDoc;

internal static class AsciiDocText {
    internal static void EnsureSingleLine(string text, string parameterName) {
        if (text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) {
            throw new ArgumentException("Value must fit on one source line.", parameterName);
        }
    }

    internal static string NormalizeLineEndings(string text, string lineEnding) {
        if (text.Length == 0) return text;
        var builder = new StringBuilder(text.Length);
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (current == '\r') {
                if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                builder.Append(lineEnding);
            } else if (current == '\n') {
                builder.Append(lineEnding);
            } else {
                builder.Append(current);
            }
        }
        return builder.ToString();
    }

    internal static bool EndsWithLineEnding(string text) =>
        text.Length > 0 && (text[text.Length - 1] == '\r' || text[text.Length - 1] == '\n');

    internal static bool IsMacroName(string value) {
        if (value.Length == 0 || !IsAsciiLetter(value[0])) return false;
        for (int index = 1; index < value.Length; index++) {
            char current = value[index];
            if (!IsAsciiLetter(current) && !char.IsDigit(current) && current != '_' && current != '-') return false;
        }
        return true;
    }

    internal static bool IsAttributeName(string value) {
        if (value.Length == 0) return false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (char.IsWhiteSpace(current) || current == ':' || current == '!') return false;
        }
        return true;
    }

    internal static bool IsAsciiLetter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z');
}
