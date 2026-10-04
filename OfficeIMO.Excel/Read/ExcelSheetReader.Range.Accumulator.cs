#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>Shares explicit and implicit row/cell coordinate handling across worksheet scans.</summary>
        private sealed class WorksheetRangeAccumulator {
            private int _minRow = int.MaxValue;
            private int _minColumn = int.MaxValue;
            private int _maxRow;
            private int _maxColumn;
            private int _nextRowIndex = 1;
            private int _rowIndex;
            private bool _hasExplicitRowIndex;
            private int _rowMinRow;
            private int _rowMaxRow;
            private int _rowMinColumn;
            private int _rowMaxColumn;
            private int _nextColumnIndex;

            internal void BeginRow(ReadOnlySpan<char> reference) {
                _rowIndex = ParsePositiveIntAttribute(reference);
                _hasExplicitRowIndex = _rowIndex > 0;
                if (!_hasExplicitRowIndex) _rowIndex = _nextRowIndex;
                _nextRowIndex = _rowIndex + 1;
                _rowMinRow = _hasExplicitRowIndex ? _rowIndex : int.MaxValue;
                _rowMaxRow = _hasExplicitRowIndex ? _rowIndex : 0;
                _rowMinColumn = int.MaxValue;
                _rowMaxColumn = 0;
                _nextColumnIndex = 1;
            }

            internal void AddCell(ReadOnlySpan<char> reference) {
                int column = 0;
                if (_hasExplicitRowIndex) {
                    column = GetXmlCellColumnIndex(reference, ref _nextColumnIndex);
                } else if (A1.TryParseCellReferenceFast(reference, out int parsedRow, out int parsedColumn)) {
                    column = parsedColumn;
                    if (parsedRow > 0) {
                        if (parsedRow < _rowMinRow) _rowMinRow = parsedRow;
                        if (parsedRow > _rowMaxRow) _rowMaxRow = parsedRow;
                    }
                    _nextColumnIndex = parsedColumn + 1;
                }
                if (column <= 0) {
                    column = _nextColumnIndex;
                    _nextColumnIndex = column + 1;
                }
                if (column > 0) {
                    if (column < _rowMinColumn) _rowMinColumn = column;
                    if (column > _rowMaxColumn) _rowMaxColumn = column;
                }
            }

            internal void EndRow() {
                if (_rowMaxColumn <= 0) return;
                if (_rowMaxRow <= 0) {
                    _rowMinRow = _rowIndex;
                    _rowMaxRow = _rowIndex;
                }
                if (_rowMinRow < _minRow) _minRow = _rowMinRow;
                if (_rowMaxRow > _maxRow) _maxRow = _rowMaxRow;
                if (_rowMinColumn < _minColumn) _minColumn = _rowMinColumn;
                if (_rowMaxColumn > _maxColumn) _maxColumn = _rowMaxColumn;
                if (!_hasExplicitRowIndex) _nextRowIndex = _rowMaxRow + 1;
            }

            internal bool TryGetReference(out string reference) {
                reference = string.Empty;
                if (_maxRow <= 0 || _maxColumn <= 0) return false;
                reference = A1.CellReference(_minRow, _minColumn) + ":" + A1.CellReference(_maxRow, _maxColumn);
                return true;
            }
        }
    }
}
