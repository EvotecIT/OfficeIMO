using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
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

        /// <summary>
        /// Performs an action on the specified cell using 1-based row and column indices.
        /// </summary>
        public TableBuilder Cell(int row, int column, Action<WordTableCell> action) {
            EnsureTable(row);
            if (_table == null) {
                return this;
            }

            while (_table.Rows.Count < row) {
                _table.AddRow(_columns);
            }

            var wordRow = _table.Rows[row - 1];
            if (column > wordRow.CellsCount) {
                for (int i = wordRow.CellsCount; i < column; i++) {
                    new WordTableCell(_fluent.Document, _table, wordRow);
                }
            }

            action(wordRow.Cells[column - 1]);
            return this;
        }

        /// <summary>
        /// Populates all cells using the provided text factory.
        /// </summary>
        public TableBuilder ForEachCell(Func<int, int, string> textFactory) {
            if (_table != null) {
                for (int r = 0; r < _table.Rows.Count; r++) {
                    var row = _table.Rows[r];
                    for (int c = 0; c < row.CellsCount; c++) {
                        row.Cells[c].AddParagraph(textFactory(r + 1, c + 1), true);
                    }
                }
            }
            return this;
        }

        /// <summary>
        /// Executes an action for each cell in the table.
        /// </summary>
        public TableBuilder ForEachCell(Action<int, int, WordTableCell> action) {
            if (_table != null) {
                for (int r = 0; r < _table.Rows.Count; r++) {
                    var row = _table.Rows[r];
                    for (int c = 0; c < row.CellsCount; c++) {
                        action(r + 1, c + 1, row.Cells[c]);
                    }
                }
            }
            return this;
        }

        /// <summary>
        /// Inserts a row at the specified 1-based index.
        /// </summary>
        public TableBuilder InsertRow(int index, params object[] cells) {
            EnsureTable();
            if (_table == null) {
                return this;
            }

            var row = _table.AddRow(_columns);
            if (index - 1 < _table._table.Elements<TableRow>().Count() - 1) {
                row._tableRow.Remove();
                _table._table.InsertAt(row._tableRow, index - 1);
            }

            for (int i = 0; i < _columns && i < cells.Length; i++) {
                row.Cells[i].AddParagraph(cells[i]?.ToString() ?? string.Empty, true);
            }

            return this;
        }

        /// <summary>
        /// Inserts a column at the specified 1-based index.
        /// </summary>
        public TableBuilder InsertColumn(int index, params object[] cells) {
            EnsureTable();
            if (_table == null) {
                return this;
            }

            int rowCount = _table.Rows.Count;
            for (int r = 0; r < rowCount; r++) {
                var row = _table.Rows[r];
                var tableCell = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }), new Paragraph());
                if (index - 1 < row.CellsCount) {
                    row._tableRow.InsertAt(tableCell, index - 1);
                } else {
                    row._tableRow.Append(tableCell);
                }
                var wordCell = new WordTableCell(_fluent.Document, _table, row, tableCell);
                if (r < cells.Length) {
                    wordCell.AddParagraph(cells[r]?.ToString() ?? string.Empty, true);
                }
            }

            _columns++;
            return this;
        }

        /// <summary>
        /// Deletes the row at the specified 1-based index.
        /// </summary>
        public TableBuilder DeleteRow(int index) {
            if (_table != null && index >= 1 && index <= _table.Rows.Count) {
                _table.Rows[index - 1].Remove();
            }
            return this;
        }

        /// <summary>
        /// Deletes the column at the specified 1-based index.
        /// </summary>
        public TableBuilder DeleteColumn(int index) {
            if (_table != null && index >= 1) {
                foreach (var row in _table.Rows) {
                    if (index <= row.CellsCount) {
                        row.Cells[index - 1]._tableCell.Remove();
                    }
                }
                if (index <= _columns) {
                    _columns--;
                }
            }
            return this;
        }

        /// <summary>
        /// Merges a rectangular range of cells using 1-based coordinates.
        /// </summary>
        public TableBuilder Merge(int fromRow, int fromColumn, int toRow, int toColumn) {
            if (_table != null) {
                int rowSpan = toRow - fromRow + 1;
                int colSpan = toColumn - fromColumn + 1;
                _table.MergeCells(fromRow - 1, fromColumn - 1, rowSpan, colSpan);
            }
            return this;
        }

        /// <summary>
        /// Applies an action to a specified row.
        /// </summary>
        public TableBuilder RowStyle(int index, Action<WordTableRow> action) {
            if (_table != null && index >= 1 && index <= _table.Rows.Count) {
                action(_table.Rows[index - 1]);
            }
            return this;
        }

        /// <summary>
        /// Applies an action to each cell in a specified column.
        /// </summary>
        public TableBuilder ColumnStyle(int index, Action<WordTableCell> action) {
            if (_table != null && index >= 1) {
                foreach (var row in _table.Rows) {
                    if (index <= row.CellsCount) {
                        action(row.Cells[index - 1]);
                    }
                }
            }
            return this;
        }
    }
}

