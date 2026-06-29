namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents one changed-cell value in a BIFF8 worksheet scenario.
    /// </summary>
    public sealed class LegacyXlsScenarioInputCell {
        /// <summary>
        /// Creates scenario changed-cell metadata.
        /// </summary>
        public LegacyXlsScenarioInputCell(string cellReference, int row, int column, bool deleted, string value) {
            CellReference = cellReference ?? throw new ArgumentNullException(nameof(cellReference));
            Row = row;
            Column = column;
            Deleted = deleted;
            Value = value ?? string.Empty;
        }

        /// <summary>Gets the changed-cell reference in A1 notation.</summary>
        public string CellReference { get; }

        /// <summary>Gets the changed-cell row as a one-based index.</summary>
        public int Row { get; }

        /// <summary>Gets the changed-cell column as a one-based index.</summary>
        public int Column { get; }

        /// <summary>Gets whether the changed cell was deleted.</summary>
        public bool Deleted { get; }

        /// <summary>Gets the scenario value text stored for the changed cell.</summary>
        public string Value { get; }
    }
}
