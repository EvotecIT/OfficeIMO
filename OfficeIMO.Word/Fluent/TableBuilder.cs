using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.Linq;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for tables.
    /// </summary>
    public class TableBuilder {
        private readonly WordFluentDocument _fluent;
        private WordTable? _table;
        private int _columns;
        private int? _preferredWidthPct;
        private int? _preferredWidthDxa;

        internal TableBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal TableBuilder(WordFluentDocument fluent, WordTable table) {
            _fluent = fluent;
            _table = table;
            _columns = table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0;
        }

        public WordTable? Table => _table;

        /// <summary>
        /// Creates the table with the specified size.
        /// </summary>
        /// <param name="rows">Number of rows.</param>
        /// <param name="columns">Number of columns.</param>
        /// <returns>The current <see cref="TableBuilder"/>.</returns>
        public TableBuilder AddTable(int rows, int columns) {
            _columns = columns;
            _table = _fluent.Document.AddTable(rows, columns);
            if (_preferredWidthPct.HasValue) {
                _table.WidthType = TableWidthUnitValues.Pct;
                _table.Width = _preferredWidthPct.Value * 50;
            } else if (_preferredWidthDxa.HasValue) {
                _table.WidthType = TableWidthUnitValues.Dxa;
                _table.Width = _preferredWidthDxa.Value;
            }
            return this;
        }

        /// <summary>
        /// Sets the number of columns for the table.
        /// </summary>
        public TableBuilder Columns(int columns) {
            _columns = columns;
            return this;
        }

        private void EnsureTable(int rows = 1) {
            if (_table == null) {
                if (_columns == 0) {
                    _columns = 1;
                }
                _table = _fluent.Document.AddTable(rows, _columns);
                if (_preferredWidthPct.HasValue) {
                    _table.WidthType = TableWidthUnitValues.Pct;
                    _table.Width = _preferredWidthPct.Value * 50;
                } else if (_preferredWidthDxa.HasValue) {
                    _table.WidthType = TableWidthUnitValues.Dxa;
                    _table.Width = _preferredWidthDxa.Value;
                }
            }
        }

        /// <summary>
        /// Sets the preferred width of the table.
        /// </summary>
        public TableBuilder PreferredWidth(int? Percent = null, int? Dxa = null) {
            if (_table != null) {
                if (Percent.HasValue) {
                    _table.WidthType = TableWidthUnitValues.Pct;
                    _table.Width = Percent.Value * 50;
                } else if (Dxa.HasValue) {
                    _table.WidthType = TableWidthUnitValues.Dxa;
                    _table.Width = Dxa.Value;
                }
            } else {
                _preferredWidthPct = Percent;
                _preferredWidthDxa = Dxa;
            }
            return this;
        }

        /// <summary>
        /// Adds a header row to the table.
        /// </summary>
        public TableBuilder Header(params object[] cells) {
            if (_columns == 0) {
                _columns = cells.Length;
            }
            EnsureTable(1);
            var row = _table!.Rows[0];
            for (int i = 0; i < _columns && i < cells.Length; i++) {
                row.Cells[i].AddParagraph(cells[i]?.ToString() ?? string.Empty, true);
            }
            _table.ConditionalFormattingFirstRow = true;
            row.RepeatHeaderRowAtTheTopOfEachPage = true;
            return this;
        }

        /// <summary>
        /// Adds a row to the table.
        /// </summary>
        public TableBuilder Row(params object[] cells) {
            if (_columns == 0) {
                _columns = cells.Length;
            }
            EnsureTable(1);
            WordTableRow row;
            if (_table!.Rows.Count == 1 && _table.Rows[0].Cells.All(c => c.Paragraphs.Count == 0 || string.IsNullOrEmpty(c.Paragraphs[0].Text))) {
                row = _table.Rows[0];
            } else {
                row = _table.AddRow(_columns);
            }
            for (int i = 0; i < _columns && i < cells.Length; i++) {
                row.Cells[i].AddParagraph(cells[i]?.ToString() ?? string.Empty, true);
            }
            return this;
        }

        /// <summary>
        /// Applies a built-in table style.
        /// </summary>
        public TableBuilder Style(WordTableStyle style) {
            if (_table != null) {
                _table.Style = style;
            }
            return this;
        }

        /// <summary>
        /// Sets horizontal alignment for the table.
        /// </summary>
        public TableBuilder Align(WordHorizontalAlignmentValues alignment) {
            if (_table != null) {
                _table.Alignment = alignment switch {
                    WordHorizontalAlignmentValues.Center => TableRowAlignmentValues.Center,
                    WordHorizontalAlignmentValues.Right => TableRowAlignmentValues.Right,
                    _ => TableRowAlignmentValues.Left,
                };
            }
            return this;
        }

        /// <summary>
        /// Creates the table from a two-dimensional array.
        /// </summary>
        public TableBuilder From2D(object[,] data) {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            _columns = cols;
            _table = _fluent.Document.AddTable(rows, cols);
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    _table.Rows[r].Cells[c].AddParagraph(data[r, c]?.ToString() ?? string.Empty, true);
                }
            }
            return this;
        }

        /// <summary>
        /// Marks a specified row as the header row.
        /// </summary>
        public TableBuilder HeaderRow(int index) {
            if (_table != null && index >= 0 && index < _table.Rows.Count) {
                _table.Rows[index].RepeatHeaderRowAtTheTopOfEachPage = true;
                _table.ConditionalFormattingFirstRow = true;
            }
            return this;
        }
    }
}

