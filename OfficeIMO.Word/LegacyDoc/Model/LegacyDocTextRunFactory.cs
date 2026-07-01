namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocTextRunFactory {
        internal static LegacyDocTextRun CreatePageNumberRun(LegacyDocCharacterFormat format, IReadOnlyList<int> characterPositions) {
            return new LegacyDocTextRun(
                string.Empty,
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
                isPageNumber: true);
        }
    }
}
