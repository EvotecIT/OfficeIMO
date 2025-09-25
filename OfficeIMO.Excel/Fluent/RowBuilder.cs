namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent helper for writing a single row.
    /// </summary>
    public class RowBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _rowIndex;

        internal RowBuilder(SheetBuilder sheetBuilder, ExcelSheet sheet, int rowIndex) {
            _sheet = sheet;
            _rowIndex = rowIndex;
        }

        /// <summary>Writes a cell in this row by 1‑based column index.</summary>
        public RowBuilder Cell(int column, object? value = null, string? formula = null, string? numberFormat = null) {
            if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
            _sheet.Cell(_rowIndex, column, value, formula, numberFormat);
            return this;
        }

        /// <summary>Writes a cell in this row using a column letter reference (e.g., "A", "BC").</summary>
        public RowBuilder Cell(string columnReference, object? value = null, string? formula = null, string? numberFormat = null) {
            if (string.IsNullOrWhiteSpace(columnReference)) throw new ArgumentNullException(nameof(columnReference));
            int column = ColumnLetterToIndex(columnReference);
            return Cell(column, value, formula, numberFormat);
        }

        /// <summary>Writes a full set of values across the row starting at column 1.</summary>
        public RowBuilder Values(params object?[] values) {
            if (values == null || values.Length == 0) return this;
            var cells = new System.Collections.Generic.List<(int Row, int Column, object Value)>(values.Length);
            for (int i = 0; i < values.Length; i++) {
                cells.Add((_rowIndex, i + 1, values[i] ?? string.Empty));
            }
            _sheet.CellValues(cells);
            return this;
        }

        private static int ColumnLetterToIndex(string column) {
            int result = 0;
            foreach (char c in column.ToUpperInvariant()) {
                if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column reference", nameof(column));
                result = result * 26 + (c - 'A' + 1);
            }
            return result;
        }
    }
}
