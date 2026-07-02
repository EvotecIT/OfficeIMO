namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocSection {
        internal LegacyDocSection(int startCharacter, int endCharacter, LegacyDocSectionFormat format) {
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Format = format;
        }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal LegacyDocSectionFormat Format { get; }
    }
}
