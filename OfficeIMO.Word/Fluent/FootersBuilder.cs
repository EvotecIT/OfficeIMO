using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for footers.
    /// </summary>
    public class FootersBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph? _paragraph;

        internal FootersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal FootersBuilder(WordFluentDocument fluent, WordParagraph? paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph? Paragraph => _paragraph;

        public FootersBuilder AddFooter(string text) {
            var footer = _fluent.Document.Footer;
            var paragraph = footer?.Default?.AddParagraph(text);
            return new FootersBuilder(_fluent, paragraph);
        }
    }
}
