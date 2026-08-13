namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool ContainsPotentialCssIdentifier(string? value, string identifier) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        for (int start = 0; start < value!.Length; start++) {
            int index = start;
            bool matches = true;
            for (int expectedIndex = 0; expectedIndex < identifier.Length; expectedIndex++) {
                if (!TryReadCssCharacter(value, ref index, out char actual)
                    || char.ToLowerInvariant(actual) != char.ToLowerInvariant(identifier[expectedIndex])) {
                    matches = false;
                    break;
                }
            }
            if (matches && index < value.Length && value[index] == '(') return true;
        }
        return false;
    }

    private static bool TryReadCssCharacter(string value, ref int index, out char character) {
        character = default;
        if (index >= value.Length) return false;
        char current = value[index++];
        if (current != '\\') {
            character = current;
            return true;
        }
        if (index >= value.Length || value[index] is '\r' or '\n' or '\f') return false;
        if (!TryGetCssHexValue(value[index], out int digit)) {
            character = value[index++];
            return true;
        }

        int scalar = 0;
        int digits = 0;
        do {
            scalar = scalar * 16 + digit;
            index++;
            digits++;
        } while (digits < 6 && index < value.Length && TryGetCssHexValue(value[index], out digit));
        if (index < value.Length && char.IsWhiteSpace(value[index])) {
            if (value[index++] == '\r' && index < value.Length && value[index] == '\n') index++;
        }
        character = scalar is > 0 and <= char.MaxValue ? (char) scalar : '\uFFFD';
        return true;
    }

    private static bool TryGetCssHexValue(char value, out int digit) {
        if (value is >= '0' and <= '9') {
            digit = value - '0';
            return true;
        }
        char lower = char.ToLowerInvariant(value);
        if (lower is >= 'a' and <= 'f') {
            digit = lower - 'a' + 10;
            return true;
        }
        digit = 0;
        return false;
    }
}
