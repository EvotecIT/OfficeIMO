namespace OfficeIMO.Excel {
    /// <summary>
    /// A typed cell value with row/column coordinates.
    /// </summary>
    public readonly struct CellValueInfo {
        /// <summary>One-based row index.</summary>
        public int Row { get; }
        /// <summary>One-based column index.</summary>
        public int Column { get; }
        /// <summary>Cell value after conversion.</summary>
        public object? Value { get; }

        /// <summary>
        /// Creates a typed cell value with coordinates.
        /// </summary>
        public CellValueInfo(int row, int column, object? value) {
            Row = row;
            Column = column;
            Value = value;
        }
    }
}
