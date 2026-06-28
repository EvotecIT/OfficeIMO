namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a BIFF8 watched-cell reference parsed from a CellWatch record.
    /// </summary>
    public sealed class LegacyXlsCellWatch {
        /// <summary>
        /// Creates watched-cell metadata.
        /// </summary>
        public LegacyXlsCellWatch(string cellReference, int row, int column) {
            CellReference = cellReference ?? throw new ArgumentNullException(nameof(cellReference));
            Row = row;
            Column = column;
        }

        /// <summary>Gets the watched cell reference in A1 notation.</summary>
        public string CellReference { get; }

        /// <summary>Gets the watched cell row as a one-based index.</summary>
        public int Row { get; }

        /// <summary>Gets the watched cell column as a one-based index.</summary>
        public int Column { get; }
    }
}
