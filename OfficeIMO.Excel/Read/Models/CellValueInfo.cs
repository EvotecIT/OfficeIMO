namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// A typed cell value with row/column coordinates.
    /// </summary>
    public readonly struct CellValueInfo
    {
        public int Row { get; }
        public int Column { get; }
        public object? Value { get; }

        public CellValueInfo(int row, int column, object? value)
        {
            Row = row;
            Column = column;
            Value = value;
        }
    }
}
