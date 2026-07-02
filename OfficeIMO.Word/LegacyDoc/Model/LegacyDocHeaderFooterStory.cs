using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocHeaderFooterStory {
        internal LegacyDocHeaderFooterStory(int sectionIndex, bool isHeader, HeaderFooterValues type, IReadOnlyList<LegacyDocHeaderFooterParagraph> paragraphs) {
            SectionIndex = sectionIndex;
            IsHeader = isHeader;
            Type = type;
            Paragraphs = paragraphs;
        }

        internal int SectionIndex { get; }

        internal bool IsHeader { get; }

        internal HeaderFooterValues Type { get; }

        internal IReadOnlyList<LegacyDocHeaderFooterParagraph> Paragraphs { get; }
    }

    internal sealed class LegacyDocHeaderFooterParagraph {
        internal LegacyDocHeaderFooterParagraph(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format)
            : this(runs, format, GetStartCharacter(runs), GetEndCharacter(runs)) {
        }

        internal LegacyDocHeaderFooterParagraph(
            IReadOnlyList<LegacyDocTextRun> runs,
            LegacyDocParagraphFormat format,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocBookmark>? bookmarks = null) {
            Runs = runs.Count == 0
                ? Array.Empty<LegacyDocTextRun>()
                : runs.ToArray();
            Format = format;
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Bookmarks = bookmarks ?? Array.Empty<LegacyDocBookmark>();
        }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; }

        internal string Text => string.Concat(Runs.Select(run => run.Text));

        private static int GetStartCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
            foreach (LegacyDocTextRun run in runs) {
                if (run.CharacterPositions.Count > 0) {
                    return run.CharacterPositions[0];
                }
            }

            return 0;
        }

        private static int GetEndCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
            for (int index = runs.Count - 1; index >= 0; index--) {
                IReadOnlyList<int> positions = runs[index].CharacterPositions;
                if (positions.Count > 0) {
                    return positions[positions.Count - 1] + 1;
                }
            }

            return 0;
        }
    }
}
