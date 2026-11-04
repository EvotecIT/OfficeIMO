namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a worksheet written or imported from tabular data.
    /// </summary>
    public sealed class ExcelDataSetImportResult {
        internal ExcelDataSetImportResult(string sheetName, string? tableName, string range, int rowCount, int columnCount) {
            SheetName = sheetName;
            TableName = tableName;
            Range = range;
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        /// <summary>Worksheet name used for the imported table.</summary>
        public string SheetName { get; }

        /// <summary>Actual Excel table name, when a table was created.</summary>
        public string? TableName { get; }

        /// <summary>A1 range occupied by the imported data, or an empty string when no worksheet cells were written.</summary>
        public string Range { get; }

        /// <summary>Number of source data rows imported.</summary>
        public int RowCount { get; }

        /// <summary>Number of source columns imported.</summary>
        public int ColumnCount { get; }
    }
}
