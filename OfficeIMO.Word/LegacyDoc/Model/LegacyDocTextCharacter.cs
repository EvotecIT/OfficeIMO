namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocTextCharacter {
        internal LegacyDocTextCharacter(char character, int fileOffset) {
            Character = character;
            FileOffset = fileOffset;
        }

        internal char Character { get; }

        internal int FileOffset { get; }
    }
}
