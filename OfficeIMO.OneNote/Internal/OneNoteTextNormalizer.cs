namespace OfficeIMO.OneNote;

/// <summary>
/// Turns native OneNote/RichEdit text into displayable Unicode: field-code markers and HYPERLINK
/// instructions are removed (see <see cref="OneNoteFieldCodes"/>), layout controls become line breaks,
/// and other controls, noncharacters and unpaired surrogates become '?'.
/// Shared by the semantic projection and the visual page canvas.
/// </summary>
internal static class OneNoteTextNormalizer {
    internal static string Normalize(string? value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
        value = OneNoteFieldCodes.StripInstructions(value!);
        if (!RequiresNormalization(value)) return value;

        var builder = new System.Text.StringBuilder(value!.Length);
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

    private static bool RequiresNormalization(string value) {
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '\v' || character == '\f' || IsUnsupportedControl(character) || IsUnicodeNoncharacter(character)) return true;
            if (char.IsHighSurrogate(character)) {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) return true;
                if (IsUnicodeNoncharacter(char.ConvertToUtf32(character, value[++index]))) return true;
            } else if (char.IsLowSurrogate(character)) {
                return true;
            }
        }
        return false;
    }

    private static bool IsUnsupportedControl(char value) =>
        value != '\t' && value != '\n' && value != '\r' &&
        (value < ' ' || value >= '\u007F' && value <= '\u009F');

    private static bool IsUnicodeNoncharacter(int scalar) =>
        scalar >= 0xFDD0 && scalar <= 0xFDEF || (scalar & 0xFFFF) >= 0xFFFE;
}
