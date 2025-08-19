using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for headers.
    /// </summary>
    public class HeadersBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph? _paragraph;

        internal HeadersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal HeadersBuilder(WordFluentDocument fluent, WordParagraph? paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph? Paragraph => _paragraph;

        public HeadersBuilder AddHeader(string text) {
            var header = _fluent.Document.Header;
            var paragraph = header?.Default?.AddParagraph(text);
            return new HeadersBuilder(_fluent, paragraph);
        }
    }
}
