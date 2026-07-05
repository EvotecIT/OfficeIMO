namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocComment {
        internal LegacyDocComment(
            int referenceCharacterPosition,
            IReadOnlyList<LegacyDocNoteParagraph> paragraphs,
            string author,
            string initials) {
            ReferenceCharacterPosition = referenceCharacterPosition;
            ParagraphRuns = paragraphs.Count == 0
                ? Array.Empty<LegacyDocNoteParagraph>()
                : paragraphs.ToArray();
            Paragraphs = ParagraphRuns.Select(paragraph => paragraph.Text).ToArray();
            Author = string.IsNullOrWhiteSpace(author) ? "Legacy DOC" : author;
            Initials = string.IsNullOrWhiteSpace(initials) ? "DOC" : initials;
        }

        internal int ReferenceCharacterPosition { get; }

        internal IReadOnlyList<string> Paragraphs { get; }

        internal IReadOnlyList<LegacyDocNoteParagraph> ParagraphRuns { get; }

        internal string Author { get; }

        internal string Initials { get; }
    }
}
