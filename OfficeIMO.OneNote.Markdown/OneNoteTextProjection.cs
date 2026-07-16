namespace OfficeIMO.OneNote.Markdown;

/// <summary>Normalizes native OneNote text for safe text, Markdown, HTML, Reader, and PDF projection.</summary>
public static class OneNoteTextProjection {
    /// <summary>
    /// Converts OneNote/RichEdit layout controls into line breaks and replaces invalid Unicode
    /// controls, noncharacters, and unpaired surrogates. The typed source model is not modified.
    /// </summary>
    /// <param name="value">Native OneNote text.</param>
    /// <returns>Projection-safe Unicode text.</returns>
    public static string Normalize(string? value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        var builder = new StringBuilder(value!.Length);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '\v' || character == '\f') {
                builder.Append('\n');
                continue;
            }

            if (IsUnsupportedControl(character)) {
                builder.Append('?');
                continue;
            }

            if (char.IsHighSurrogate(character)) {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) {
                    builder.Append('?');
                    continue;
                }

                char low = value[++index];
                int scalar = char.ConvertToUtf32(character, low);
                if (IsUnicodeNoncharacter(scalar)) builder.Append('?');
                else builder.Append(character).Append(low);
                continue;
            }

            if (char.IsLowSurrogate(character) || IsUnicodeNoncharacter(character)) {
                builder.Append('?');
                continue;
            }

            builder.Append(character);
        }

        return builder.ToString();
    }

    private static bool IsUnsupportedControl(char value) =>
        value != '\t' && value != '\n' && value != '\r' &&
        (value < ' ' || value >= '\u007F' && value <= '\u009F');

    private static bool IsUnicodeNoncharacter(int scalar) =>
        scalar >= 0xFDD0 && scalar <= 0xFDEF || (scalar & 0xFFFF) >= 0xFFFE;
}
