using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent helper for operating over a rectangular A1 range (formatting, setting, clearing, per‑cell writes).
    /// </summary>
    public class RangeBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _fromRow;
        private readonly int _fromCol;
        private readonly int _toRow;
        private readonly int _toCol;

        internal RangeBuilder(ExcelSheet sheet, int fromRow, int fromCol, int toRow, int toCol) {
            _sheet = sheet;
            _fromRow = fromRow;
            _fromCol = fromCol;
            _toRow = toRow;
            _toCol = toCol;
        }

        /// <summary>Applies a number format to all cells in the range.</summary>
        public RangeBuilder NumberFormat(string numberFormat) {
            for (int r = _fromRow; r <= _toRow; r++) {
                for (int c = _fromCol; c <= _toCol; c++) {
                    _sheet.FormatCell(r, c, numberFormat);
                }
            }
            return this;
        }

        /// <summary>Applies styles to the range via a nested <see cref="StyleBuilder"/>.</summary>
        public RangeBuilder Style(Action<StyleBuilder> action) {
            var builder = new StyleBuilder(_sheet);
            action(builder);
            return this;
        }

        /// <summary>
        /// Writes a 2D array of values into the range. The array dimensions must match the range size.
        /// </summary>
        public RangeBuilder Set(object[,] values) {
            int rows = _toRow - _fromRow + 1;
            int cols = _toCol - _fromCol + 1;
            if (values.GetLength(0) != rows || values.GetLength(1) != cols) {
                throw new ArgumentException("Values array dimensions must match the range.", nameof(values));
            }

            var cells = new List<(int Row, int Column, object Value)>();
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    cells.Add((_fromRow + r, _fromCol + c, values[r, c]));
                }
            }
            _sheet.CellValuesParallel(cells);
            return this;
        }

        /// <summary>Clears all cells in the range.</summary>
        public RangeBuilder Clear() {
            for (int r = _fromRow; r <= _toRow; r++) {
                for (int c = _fromCol; c <= _toCol; c++) {
                    _sheet.CellValue(r, c, string.Empty);
                }
            }
            return this;
        }

        /// <summary>
        /// Writes an individual cell within the range by offset (1‑based) from the top‑left corner.
        /// </summary>
        public RangeBuilder Cell(int rowOffset, int columnOffset, object? value = null, string? formula = null, string? numberFormat = null) {
            if (rowOffset < 1) throw new ArgumentOutOfRangeException(nameof(rowOffset));
            if (columnOffset < 1) throw new ArgumentOutOfRangeException(nameof(columnOffset));
            int row = _fromRow + rowOffset - 1;
            int col = _fromCol + columnOffset - 1;
            if (row > _toRow) throw new ArgumentOutOfRangeException(nameof(rowOffset));
            if (col > _toCol) throw new ArgumentOutOfRangeException(nameof(columnOffset));
            _sheet.Cell(row, col, value, formula, numberFormat);
            return this;
        }
    }
}

