using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocHeaderFooterStory {
        internal LegacyDocHeaderFooterStory(int sectionIndex, bool isHeader, HeaderFooterValues type, IReadOnlyList<string> paragraphs) {
            SectionIndex = sectionIndex;
            IsHeader = isHeader;
            Type = type;
            Paragraphs = paragraphs;
        }

        internal int SectionIndex { get; }

        internal bool IsHeader { get; }

        internal HeaderFooterValues Type { get; }

        internal IReadOnlyList<string> Paragraphs { get; }
    }
}
