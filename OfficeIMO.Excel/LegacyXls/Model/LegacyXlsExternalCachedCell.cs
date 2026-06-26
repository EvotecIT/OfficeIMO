namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one preserved value from an external cell cache CRN record.
    /// </summary>
    public sealed class LegacyXlsExternalCachedCell {
        internal LegacyXlsExternalCachedCell(int row, int column, LegacyXlsCellValueKind kind, object? value) {
            Row = row;
            Column = column;
            Kind = kind;
            Value = value;
        }

        /// <summary>
        /// Gets the zero-based row index in the external sheet cache.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the zero-based column index in the external sheet cache.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the cached value category.
        /// </summary>
        public LegacyXlsCellValueKind Kind { get; }

        /// <summary>
        /// Gets the cached value. Error values are stored as Excel error text.
        /// </summary>
        public object? Value { get; }
    }
}
