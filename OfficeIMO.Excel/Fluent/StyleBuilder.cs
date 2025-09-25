namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent styling helpers for individual cells.
    /// </summary>
    public class StyleBuilder {
        private readonly ExcelSheet _sheet;

        internal StyleBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        /// <summary>
        /// Applies a number format to the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="numberFormat">The number format code to apply.</param>
        /// <returns>The current <see cref="StyleBuilder"/> instance for fluent chaining.</returns>
        public StyleBuilder FormatCell(int row, int column, string numberFormat) {
            _sheet.FormatCell(row, column, numberFormat);
            return this;
        }
    }
}
