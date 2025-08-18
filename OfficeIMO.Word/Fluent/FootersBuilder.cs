namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for footers.
    /// </summary>
    public class FootersBuilder {
        private readonly WordFluentDocument _fluent;

        internal FootersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddFooter(string text) {
            var footer = _fluent.Document.Footer;
            footer?.Default?.AddParagraph(text);
            return _fluent;
        }
    }
}
