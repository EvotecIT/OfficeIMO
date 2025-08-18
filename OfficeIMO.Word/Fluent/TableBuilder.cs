namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for tables.
    /// </summary>
    public class TableBuilder {
        private readonly WordFluentDocument _fluent;

        internal TableBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddTable(int rows, int columns) {
            _fluent.Document.AddTable(rows, columns);
            return _fluent;
        }
    }
}
