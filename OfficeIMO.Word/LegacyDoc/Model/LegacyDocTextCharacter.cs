namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocTextCharacter {
        internal LegacyDocTextCharacter(char character, int fileOffset, int characterPosition) {
            Character = character;
            FileOffset = fileOffset;
            CharacterPosition = characterPosition;
        }

        internal char Character { get; }

        internal int FileOffset { get; }

        internal int CharacterPosition { get; }
    }
}
