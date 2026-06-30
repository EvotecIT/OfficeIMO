namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocBookmark {
        internal LegacyDocBookmark(string name, int startCharacter, int endCharacter) {
            Name = name;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
        }

        internal string Name { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal bool IsZeroLength => StartCharacter == EndCharacter;
    }
}
