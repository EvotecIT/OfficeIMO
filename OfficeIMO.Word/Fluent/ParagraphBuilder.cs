namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for paragraphs.
    /// </summary>
    public class ParagraphBuilder {
        private readonly WordFluentDocument _fluent;

        internal ParagraphBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddParagraph(string text) {
            _fluent.Document.AddParagraph(text);
            return _fluent;
        }
    }
}
