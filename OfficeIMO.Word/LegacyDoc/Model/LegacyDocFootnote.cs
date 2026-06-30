namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocFootnote {
        internal LegacyDocFootnote(int referenceCharacterPosition, IReadOnlyList<string> paragraphs) {
            ReferenceCharacterPosition = referenceCharacterPosition;
            ParagraphRuns = paragraphs.Count == 0
                ? Array.Empty<LegacyDocNoteParagraph>()
                : paragraphs.Select(CreatePlainParagraph).ToArray();
            Paragraphs = ParagraphRuns.Select(paragraph => paragraph.Text).ToArray();
        }

        internal LegacyDocFootnote(int referenceCharacterPosition, IReadOnlyList<LegacyDocNoteParagraph> paragraphs) {
            ReferenceCharacterPosition = referenceCharacterPosition;
            ParagraphRuns = paragraphs.Count == 0
                ? Array.Empty<LegacyDocNoteParagraph>()
                : paragraphs.ToArray();
            Paragraphs = ParagraphRuns.Select(paragraph => paragraph.Text).ToArray();
        }

        internal int ReferenceCharacterPosition { get; }

        internal IReadOnlyList<string> Paragraphs { get; }

        internal IReadOnlyList<LegacyDocNoteParagraph> ParagraphRuns { get; }

        private static LegacyDocNoteParagraph CreatePlainParagraph(string text) {
            return new LegacyDocNoteParagraph(new[] {
                new LegacyDocTextRun(
                    text,
                    bold: false,
                    italic: false,
                    strike: false,
                    doubleStrike: false,
                    outline: false,
                    shadow: false,
                    emboss: false,
                    imprint: false,
                    hidden: false,
                    caps: null,
                    verticalPosition: null,
                    underline: null,
                    highlight: null,
                    fontSizeHalfPoints: null,
                    colorHex: null,
                    fontFamily: null)
            });
        }
    }
}
