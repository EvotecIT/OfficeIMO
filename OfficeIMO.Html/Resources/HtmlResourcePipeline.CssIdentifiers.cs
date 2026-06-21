namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static bool CssFunctionNameEquals(string rawName, string functionName) {
        string trimmed = rawName.Trim();
        if (string.Equals(DecodeCssEscapes(trimmed), functionName, StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        int cursor = 0;
        for (int expectedIndex = 0; expectedIndex < functionName.Length; expectedIndex++) {
            if (!TryConsumeCssIdentifierCharacter(trimmed, ref cursor, functionName[expectedIndex], functionName.Length - expectedIndex - 1)) {
                return false;
            }
        }

        return cursor == trimmed.Length;
    }

    private static bool TryConsumeCssIdentifierCharacter(string text, ref int cursor, char expected, int remainingExpectedCharacters) {
        if (cursor >= text.Length) {
            return false;
        }

        if (text[cursor] != '\\') {
            if (!CssIdentifierCharactersEqual(text[cursor], expected)) {
                return false;
            }

            cursor++;
            return true;
        }

        if (cursor + 1 >= text.Length) {
            return false;
        }

        int hexStart = cursor + 1;
        if (!IsHexDigit(text[hexStart])) {
            if (!CssIdentifierCharactersEqual(text[hexStart], expected)) {
                return false;
            }

            cursor = hexStart + 1;
            return true;
        }

        int hexEnd = hexStart;
        while (hexEnd < text.Length && hexEnd - hexStart < 6 && IsHexDigit(text[hexEnd])) {
            hexEnd++;
        }

        int maxHexLength = hexEnd - hexStart;
        for (int hexLength = maxHexLength; hexLength >= 1; hexLength--) {
            if (text.Length - (hexStart + hexLength) < remainingExpectedCharacters) {
                continue;
            }

            string hex = text.Substring(hexStart, hexLength);
            if (!int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int codePoint)
                || codePoint <= 0
                || codePoint > 0x10FFFF
                || (codePoint >= 0xD800 && codePoint <= 0xDFFF)) {
                continue;
            }

            string decoded = char.ConvertFromUtf32(codePoint);
            if (decoded.Length == 1 && CssIdentifierCharactersEqual(decoded[0], expected)) {
                cursor = hexStart + hexLength;
                if (cursor < text.Length && char.IsWhiteSpace(text[cursor])) {
                    cursor++;
                }

                return true;
            }
        }

        return false;
    }

    private static bool CssIdentifierCharactersEqual(char left, char right) {
        return char.ToUpperInvariant(left) == char.ToUpperInvariant(right);
    }
}
