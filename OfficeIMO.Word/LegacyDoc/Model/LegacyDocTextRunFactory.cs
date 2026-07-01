namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocTextRunFactory {
        internal static LegacyDocTextRun CreateFieldRun(string resultText, LegacyDocFieldKind fieldKind, LegacyDocCharacterFormat format, IReadOnlyList<int> characterPositions) {
            return CreateFieldRun(resultText, fieldKind, fieldInstruction: null, format, characterPositions);
        }

        internal static LegacyDocTextRun CreateFieldRun(string resultText, LegacyDocFieldKind fieldKind, string? fieldInstruction, LegacyDocCharacterFormat format, IReadOnlyList<int> characterPositions) {
            return new LegacyDocTextRun(
                resultText,
                format.Bold,
                format.Italic,
                format.Strike,
                format.DoubleStrike,
                format.Outline,
                format.Shadow,
                format.Emboss,
                format.Imprint,
                format.Hidden,
                format.NoProof,
                format.Caps,
                format.VerticalPosition,
                format.Underline,
                format.Highlight,
                format.FontSizeHalfPoints,
                format.ColorHex,
                format.FontFamily,
                characterPositions,
                fieldKind: fieldKind,
                fieldInstruction: fieldInstruction);
        }
    }
}
