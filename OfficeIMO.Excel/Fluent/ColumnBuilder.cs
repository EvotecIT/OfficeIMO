namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for adjusting a single column (width, visibility) and writing cells in that column.
    /// </summary>
    public class ColumnBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _columnIndex;

        internal ColumnBuilder(ExcelSheet sheet, int columnIndex) {
            if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            _sheet = sheet;
            _columnIndex = columnIndex;
        }

        /// <summary>Auto-fits the column to its contents.</summary>
        public ColumnBuilder AutoFit() {
            _sheet.AutoFitColumn(_columnIndex);
            return this;
        }

        /// <summary>Sets the column width (Excel column units).</summary>
        public ColumnBuilder Width(double width) {
            _sheet.SetColumnWidth(_columnIndex, width);
            return this;
        }

        /// <summary>Hides or shows the column.</summary>
        public ColumnBuilder Hidden(bool hidden) {
            _sheet.SetColumnHidden(_columnIndex, hidden);
            return this;
        }

        /// <summary>
        /// Writes a cell in this column with optional value, formula, and number format.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="value">Optional value.</param>
        /// <param name="formula">Optional A1-style formula.</param>
        /// <param name="numberFormat">Optional number format code.</param>
        public ColumnBuilder Cell(int row, object? value = null, string? formula = null, string? numberFormat = null) {
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
            _sheet.Cell(row, _columnIndex, value, formula, numberFormat);
            return this;
        }
    }
}
