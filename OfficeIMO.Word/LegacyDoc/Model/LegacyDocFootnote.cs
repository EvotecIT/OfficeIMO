namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocFootnote {
        internal LegacyDocFootnote(int referenceCharacterPosition, IReadOnlyList<string> paragraphs) {
            ReferenceCharacterPosition = referenceCharacterPosition;
            Paragraphs = paragraphs.Count == 0
                ? Array.Empty<string>()
                : paragraphs.ToArray();
        }

        internal int ReferenceCharacterPosition { get; }

        internal IReadOnlyList<string> Paragraphs { get; }
    }
}
