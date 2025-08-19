using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for tables.
    /// </summary>
    public class TableBuilder {
        private readonly WordFluentDocument _fluent;
        internal WordTable? _table;
        private int _columns;
        private double? _preferredWidthPercent;
        private double? _preferredWidthPoints;

        internal TableBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal TableBuilder(WordFluentDocument fluent, WordTable table) {
            _fluent = fluent;
            _table = table;
            _columns = table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0;
        }

        public WordTable? Table => _table;

        public TableBuilder Columns(int count) {
            _columns = count;
            return this;
        }

        public TableBuilder Create(int rows, int cols) {
            _columns = cols;
            _table = _fluent.Document.AddTable(rows, cols);
            ApplyPreferredWidth();
            return this;
        }

        private void ApplyPreferredWidth() {
            if (_table == null) return;
            if (_preferredWidthPercent.HasValue) {
                _table.WidthType = TableWidthUnitValues.Pct;
                _table.Width = (int)(_preferredWidthPercent.Value * 50);
            } else if (_preferredWidthPoints.HasValue) {
                _table.WidthType = TableWidthUnitValues.Dxa;
                _table.Width = (int)(_preferredWidthPoints.Value);
            }
        }

        private void EnsureTable(int rows = 1) {
            if (_table == null) {
                if (_columns == 0) {
                    _columns = 1;
                }
                _table = _fluent.Document.AddTable(rows, _columns);
                ApplyPreferredWidth();
            } else {
                while (_table.Rows.Count < rows) {
                    _table.AddRow(_columns);
                }
            }
        }

        public TableBuilder PreferredWidth(double? percent = null, double? points = null) {
            if (_table != null) {
                _preferredWidthPercent = percent;
                _preferredWidthPoints = points;
                ApplyPreferredWidth();
            } else {
                _preferredWidthPercent = percent;
                _preferredWidthPoints = points;
            }
            return this;
        }

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

        public TableBuilder Row(params object[] cells) {
            if (_columns == 0) {
                _columns = cells.Length;
            }
            EnsureTable((_table?.Rows.Count ?? 0) + 1);
            var rowIndex = _table!.Rows.Count - 1;
            var row = _table.Rows[rowIndex];
            for (int i = 0; i < _columns && i < cells.Length; i++) {
                row.Cells[i].AddParagraph(cells[i]?.ToString() ?? string.Empty, true);
            }
            return this;
        }

        public TableBuilder From2D(object[,] data) {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            _columns = cols;
            _table = _fluent.Document.AddTable(rows, cols);
            ApplyPreferredWidth();
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    _table.Rows[r].Cells[c].AddParagraph(data[r, c]?.ToString() ?? string.Empty, true);
                }
            }
            return this;
        }

        public TableBuilder HeaderRow(int rowIndex) {
            if (_table != null && rowIndex > 0 && rowIndex <= _table.Rows.Count) {
                var row = _table.Rows[rowIndex - 1];
                row.RepeatHeaderRowAtTheTopOfEachPage = true;
                _table.ConditionalFormattingFirstRow = true;
            }
            return this;
        }

        public TableBuilder InsertRow(int index, params object[] cells) {
            EnsureTable(index);
            if (_table == null) return this;
            var row = new WordTableRow(_fluent.Document, _table);
            for (int i = 0; i < _columns; i++) {
                new WordTableCell(_fluent.Document, _table, row);
            }
            _table._table.InsertAt(row._tableRow, index - 1);
            for (int i = 0; i < _columns && i < cells.Length; i++) {
                row.Cells[i].AddParagraph(cells[i]?.ToString() ?? string.Empty, true);
            }
            return this;
        }

        public TableBuilder DeleteRow(int index) {
            if (_table != null && index > 0 && index <= _table.Rows.Count) {
                _table.Rows[index - 1].Remove();
            }
            return this;
        }

        public TableBuilder InsertColumn(int index, params object[] cells) {
            EnsureTable();
            if (_table == null) return this;
            int rowIdx = 0;
            foreach (var row in _table.Rows) {
                var cell = new WordTableCell(_fluent.Document, _table, row);
                row._tableRow.RemoveChild(cell._tableCell);
                row._tableRow.InsertAt(cell._tableCell, index - 1);
                if (rowIdx < cells.Length) {
                    row.Cells[index - 1].AddParagraph(cells[rowIdx]?.ToString() ?? string.Empty, true);
                }
                rowIdx++;
            }
            _columns++;
            return this;
        }

        public TableBuilder DeleteColumn(int index) {
            if (_table != null) {
                foreach (var row in _table.Rows) {
                    if (index > 0 && index <= row.CellsCount) {
                        row.Cells[index - 1].Remove();
                    }
                }
                if (_columns > 0) _columns--;
            }
            return this;
        }

        public TableBuilder Merge(int fromRow, int fromCol, int toRow, int toCol) {
            if (_table != null) {
                int rowSpan = toRow - fromRow + 1;
                int colSpan = toCol - fromCol + 1;
                _table.MergeCells(fromRow - 1, fromCol - 1, rowSpan, colSpan);
            }
            return this;
        }

        public TableBuilder Style(WordTableStyle style) {
            if (_table != null) {
                _table.Style = style;
            }
            return this;
        }

        public TableBuilder Align(HorizontalAlignment align) {
            if (_table != null) {
                _table.Alignment = align switch {
                    HorizontalAlignment.Center => TableRowAlignmentValues.Center,
                    HorizontalAlignment.Right => TableRowAlignmentValues.Right,
                    _ => TableRowAlignmentValues.Left,
                };
            }
            return this;
        }

        public RowBuilder Row(int index) {
            EnsureTable(index);
            return new RowBuilder(this, index - 1);
        }

        public ColumnBuilder Column(int index) {
            EnsureTable();
            return new ColumnBuilder(this, index - 1);
        }

        public CellBuilder Cell(int row, int col) {
            EnsureTable(row);
            return new CellBuilder(this, row - 1, col - 1);
        }

        public TableBuilder Cell(int row, int col, Action<CellBuilder> action) {
            var builder = Cell(row, col);
            action(builder);
            return this;
        }

        public TableBuilder ForEachCell(Func<int, int, string> textFactory) {
            EnsureTable();
            if (_table == null) return this;
            for (int r = 0; r < _table.Rows.Count; r++) {
                for (int c = 0; c < _table.Rows[r].CellsCount; c++) {
                    var text = textFactory(r + 1, c + 1);
                    _table.Rows[r].Cells[c].AddParagraph(text ?? string.Empty, true);
                }
            }
            return this;
        }

        public TableBuilder ForEachCell(Action<int, int, CellBuilder> action) {
            EnsureTable();
            if (_table == null) return this;
            for (int r = 0; r < _table.Rows.Count; r++) {
                for (int c = 0; c < _table.Rows[r].CellsCount; c++) {
                    action(r + 1, c + 1, new CellBuilder(this, r, c));
                }
            }
            return this;
        }
    }

    public sealed class RowBuilder {
        private readonly TableBuilder _parent;
        private readonly int _index;

        internal RowBuilder(TableBuilder parent, int index) {
            _parent = parent;
            _index = index;
        }

        private WordTableRow Row => _parent._table!.Rows[_index];

        public TableBuilder EachCell(Action<CellBuilder> cell) {
            for (int c = 0; c < Row.CellsCount; c++) {
                cell(new CellBuilder(_parent, _index, c));
            }
            return _parent;
        }

        public TableBuilder Shading(string hex) {
            foreach (var c in Row.Cells) {
                c.ShadingFillColorHex = hex.TrimStart('#');
            }
            return _parent;
        }

        public TableBuilder Align(HorizontalAlignment align) {
            for (int c = 0; c < Row.CellsCount; c++) {
                new CellBuilder(_parent, _index, c).Align(align);
            }
            return _parent;
        }
    }

    public sealed class ColumnBuilder {
        private readonly TableBuilder _parent;
        private readonly int _index;

        internal ColumnBuilder(TableBuilder parent, int index) {
            _parent = parent;
            _index = index;
        }

        public TableBuilder Shading(string hex) {
            foreach (var row in _parent._table!.Rows) {
                if (_index < row.CellsCount) {
                    row.Cells[_index].ShadingFillColorHex = hex.TrimStart('#');
                }
            }
            return _parent;
        }

        public TableBuilder Align(HorizontalAlignment align) {
            for (int r = 0; r < _parent._table!.Rows.Count; r++) {
                if (_index < _parent._table.Rows[r].CellsCount) {
                    new CellBuilder(_parent, r, _index).Align(align);
                }
            }
            return _parent;
        }
    }

    public sealed class CellBuilder {
        private readonly TableBuilder _parent;
        private readonly WordTableCell _cell;

        internal CellBuilder(TableBuilder parent, int rowIndex, int colIndex) {
            _parent = parent;
            _cell = parent._table!.Rows[rowIndex].Cells[colIndex];
        }

        public TableBuilder Text(string text) {
            _cell.AddParagraph(text, true);
            return _parent;
        }

        public TableBuilder Shading(string hex) {
            _cell.ShadingFillColorHex = hex.TrimStart('#');
            return _parent;
        }

        public TableBuilder Align(HorizontalAlignment hAlign, VerticalAlignment vAlign = VerticalAlignment.Center) {
            var paragraph = _cell.Paragraphs.FirstOrDefault() ?? _cell.AddParagraph("", true);
            paragraph.SetAlignment(hAlign switch {
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                _ => JustificationValues.Left,
            });
            _cell.VerticalAlignment = vAlign switch {
                VerticalAlignment.Top => TableVerticalAlignmentValues.Top,
                VerticalAlignment.Bottom => TableVerticalAlignmentValues.Bottom,
                _ => TableVerticalAlignmentValues.Center,
            };
            return _parent;
        }

        public TableBuilder Padding(double top, double right, double bottom, double left) {
            var tcPr = _cell._tableCellProperties ?? (_cell._tableCellProperties = new TableCellProperties());
            var margin = tcPr.GetFirstChild<TableCellMargin>() ?? tcPr.AppendChild(new TableCellMargin());
            margin.TopMargin = new TopMargin { Width = ((int)(top * 20)).ToString(), Type = TableWidthUnitValues.Dxa };
            margin.RightMargin = new RightMargin { Width = ((int)(right * 20)).ToString(), Type = TableWidthUnitValues.Dxa };
            margin.BottomMargin = new BottomMargin { Width = ((int)(bottom * 20)).ToString(), Type = TableWidthUnitValues.Dxa };
            margin.LeftMargin = new LeftMargin { Width = ((int)(left * 20)).ToString(), Type = TableWidthUnitValues.Dxa };
            return _parent;
        }

        public TableBuilder MergeRight(int count = 1) {
            _cell.MergeHorizontally(count, false);
            return _parent;
        }

        public TableBuilder MergeDown(int count = 1) {
            _cell.MergeVertically(count, false);
            return _parent;
        }
    }
}
