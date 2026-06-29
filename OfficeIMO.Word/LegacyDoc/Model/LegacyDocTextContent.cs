namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextContent {
        internal LegacyDocTextContent(string text, IReadOnlyList<LegacyDocTextCharacter> characters) {
            Text = text;
            Characters = characters;
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextCharacter> Characters { get; }
    }
}
