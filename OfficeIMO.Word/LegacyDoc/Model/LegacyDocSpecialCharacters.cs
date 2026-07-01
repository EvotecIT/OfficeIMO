namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocSpecialCharacters {
        internal const char TextWrappingBreak = '\v';
        internal const char PageBreak = '\f';
        internal const char ColumnBreak = '\u000E';

        internal static bool IsSupportedInlineControl(char character) {
            return character == '\t'
                || character == TextWrappingBreak
                || character == PageBreak
                || character == ColumnBreak;
        }
    }
}
