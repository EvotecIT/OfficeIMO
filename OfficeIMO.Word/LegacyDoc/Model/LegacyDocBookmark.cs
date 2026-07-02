namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocBookmark {
        internal LegacyDocBookmark(string name, int startCharacter, int endCharacter, string projectionId) {
            Name = name;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            ProjectionId = projectionId;
        }

        internal string Name { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal string ProjectionId { get; }

        internal bool IsZeroLength => StartCharacter == EndCharacter;
    }
}
