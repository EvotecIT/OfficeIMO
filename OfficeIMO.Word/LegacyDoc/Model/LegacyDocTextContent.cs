namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextContent {
        internal LegacyDocTextContent(string text, IReadOnlyList<LegacyDocTextCharacter> characters)
            : this(text, characters, characters) {
        }

        internal LegacyDocTextContent(string text, IReadOnlyList<LegacyDocTextCharacter> characters, IReadOnlyList<LegacyDocTextCharacter> allCharacters) {
            Text = text;
            Characters = characters;
            AllCharacters = allCharacters;
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextCharacter> Characters { get; }

        internal IReadOnlyList<LegacyDocTextCharacter> AllCharacters { get; }
    }
}
