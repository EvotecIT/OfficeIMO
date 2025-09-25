namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent helper for configuring one or more columns (widths, styles) by index.
    /// </summary>
    public class ColumnCollectionBuilder {
        private readonly ExcelSheet _sheet;

        internal ColumnCollectionBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        /// <summary>
        /// Configures a single column by 1‑based index via the provided action.
        /// </summary>
        public ColumnCollectionBuilder Col(int index, Action<ColumnBuilder> action) {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
            var builder = new ColumnBuilder(_sheet, index);
            action(builder);
            return this;
        }
    }
}
