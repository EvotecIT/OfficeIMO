namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a table defined in an Excel workbook.
    /// </summary>
    public sealed class ExcelTableInfo {
        /// <summary>Initializes a new instance of the <see cref="ExcelTableInfo"/> class.</summary>
        /// <param name="name">Table name (or display name).</param>
        /// <param name="range">Table range in A1 notation.</param>
        /// <param name="sheetName">Sheet name containing the table.</param>
        /// <param name="sheetIndex">0-based sheet index; -1 when unknown.</param>
        public ExcelTableInfo(string name, string range, string sheetName, int sheetIndex) {
            Name = name ?? string.Empty;
            Range = range ?? string.Empty;
            SheetName = sheetName ?? string.Empty;
            SheetIndex = sheetIndex;
        }

        /// <summary>Table name (or display name).</summary>
        public string Name { get; }

        /// <summary>Table range in A1 notation.</summary>
        public string Range { get; }

        /// <summary>Sheet name containing the table.</summary>
        public string SheetName { get; }

        /// <summary>0-based sheet index; -1 when unknown.</summary>
        public int SheetIndex { get; }
    }
}
