using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for paragraphs.
    /// </summary>
    public class ParagraphBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph? _paragraph;

        internal ParagraphBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal ParagraphBuilder(WordFluentDocument fluent, WordParagraph paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph? Paragraph => _paragraph;

        public WordFluentDocument AddParagraph(string text) {
            _fluent.Document.AddParagraph(text);
            return _fluent;
        }
    }
}
