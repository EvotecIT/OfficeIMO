using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    public class StyleMissingEventArgs : EventArgs {
        public StyleMissingEventArgs(WordParagraph paragraph, string className) {
            Paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            ClassName = className ?? throw new ArgumentNullException(nameof(className));
        }

        public WordParagraph Paragraph { get; }
        public string ClassName { get; }
        public WordParagraphStyles? Style { get; set; }
        public string? StyleId { get; set; }
    }
}
