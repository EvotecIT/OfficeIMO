namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for document information such as properties.
    /// </summary>
    public class InfoBuilder {
        private readonly WordFluentDocument _fluent;

        internal InfoBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument SetTitle(string title) {
            _fluent.Document.BuiltinDocumentProperties.Title = title;
            return _fluent;
        }
    }
}
