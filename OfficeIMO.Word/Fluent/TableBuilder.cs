using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for tables.
    /// </summary>
    public class TableBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordTable? _table;

        internal TableBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal TableBuilder(WordFluentDocument fluent, WordTable table) {
            _fluent = fluent;
            _table = table;
        }

        public WordTable? Table => _table;

        public TableBuilder AddTable(int rows, int columns) {
            var table = _fluent.Document.AddTable(rows, columns);
            return new TableBuilder(_fluent, table);
        }
    }
}
