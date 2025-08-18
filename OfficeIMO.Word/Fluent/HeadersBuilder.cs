namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for headers.
    /// </summary>
    public class HeadersBuilder {
        private readonly WordFluentDocument _fluent;

        internal HeadersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddHeader(string text) {
            var header = _fluent.Document.Header;
            header?.Default?.AddParagraph(text);
            return _fluent;
        }
    }
}
