using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    internal static class LegacyPptXmlText {
        internal static string? SanitizeAttributeValue(string? value) {
            if (value == null || value.Length == 0) return value;
            StringBuilder? sanitized = null;
            for (int index = 0; index < value.Length; index++) {
                char current = value[index];
                if (IsValidBmpCharacter(current)) {
                    sanitized?.Append(current);
                    continue;
                }
                if (char.IsHighSurrogate(current)
                    && index + 1 < value.Length
                    && char.IsLowSurrogate(value[index + 1])) {
                    if (sanitized != null) {
                        sanitized.Append(current);
                        sanitized.Append(value[index + 1]);
                    }
                    index++;
                    continue;
                }
                if (sanitized == null) {
                    sanitized = new StringBuilder(value.Length);
                    sanitized.Append(value, 0, index);
                }
            }
            return sanitized?.ToString() ?? value;
        }

        internal static bool IsValidAttributeCharacter(char value) =>
            IsValidBmpCharacter(value);

        private static bool IsValidBmpCharacter(char value) =>
            value == '\t' || value == '\n' || value == '\r'
            || value >= ' ' && value <= '\uD7FF'
            || value >= '\uE000' && value <= '\uFFFD';
    }
}
