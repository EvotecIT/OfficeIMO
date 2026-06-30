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
        internal LegacyDocHeaderFooterParagraph(IReadOnlyList<LegacyDocTextRun> runs) {
            Runs = runs.Count == 0
                ? Array.Empty<LegacyDocTextRun>()
                : runs.ToArray();
        }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal string Text => string.Concat(Runs.Select(run => run.Text));
    }
}
