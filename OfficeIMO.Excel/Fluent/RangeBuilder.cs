using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Fluent {
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

        public RangeBuilder NumberFormat(string numberFormat) {
            for (int r = _fromRow; r <= _toRow; r++) {
                for (int c = _fromCol; c <= _toCol; c++) {
                    _sheet.FormatCell(r, c, numberFormat);
                }
            }
            return this;
        }

        public RangeBuilder Style(Action<StyleBuilder> action) {
            var builder = new StyleBuilder(_sheet);
            action(builder);
            return this;
        }

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

        public RangeBuilder Clear() {
            for (int r = _fromRow; r <= _toRow; r++) {
                for (int c = _fromCol; c <= _toCol; c++) {
                    _sheet.CellValue(r, c, string.Empty);
                }
            }
            return this;
        }
    }
}

