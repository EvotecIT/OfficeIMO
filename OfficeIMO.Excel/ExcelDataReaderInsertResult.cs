namespace OfficeIMO.Excel {
    /// <summary>Describes rows inserted into a worksheet from an <see cref="System.Data.IDataReader"/>.</summary>
    public sealed class ExcelDataReaderInsertResult {
        internal ExcelDataReaderInsertResult(string sheetName, string? tableName, string range, int rowCount, int columnCount) {
            SheetName = sheetName;
            TableName = tableName;
            Range = range;
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        /// <summary>Gets the worksheet containing the inserted rows.</summary>
        public string SheetName { get; }

        /// <summary>Gets the actual Excel table name, when a table was created.</summary>
        public string? TableName { get; }

        /// <summary>Gets the occupied A1 range, or an empty string when no cells were written.</summary>
        public string Range { get; }

        /// <summary>Gets the number of source data rows inserted.</summary>
        public int RowCount { get; }

        /// <summary>Gets the number of source columns inserted.</summary>
        public int ColumnCount { get; }
    }
}
