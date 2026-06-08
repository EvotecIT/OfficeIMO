namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Describes one logical PDF table imported into an Excel worksheet.
    /// </summary>
    public sealed class PdfExcelTableImportResult {
        internal PdfExcelTableImportResult(
            int pageIndex,
            int pageNumber,
            int tableIndex,
            string detectionKind,
            string sheetName,
            string tableName,
            string range,
            int columnCount,
            int rowCount,
            int totalRowCount,
            bool truncated) {
            PageIndex = pageIndex;
            PageNumber = pageNumber;
            TableIndex = tableIndex;
            DetectionKind = detectionKind ?? string.Empty;
            SheetName = sheetName ?? string.Empty;
            TableName = tableName ?? string.Empty;
            Range = range ?? string.Empty;
            ColumnCount = columnCount;
            RowCount = rowCount;
            TotalRowCount = totalRowCount;
            Truncated = truncated;
        }

        /// <summary>Zero-based page index within the selected logical page collection.</summary>
        public int PageIndex { get; }

        /// <summary>One-based source page number from the PDF document.</summary>
        public int PageNumber { get; }

        /// <summary>Zero-based table index within the source logical PDF page.</summary>
        public int TableIndex { get; }

        /// <summary>Detection heuristic that produced the imported table.</summary>
        public string DetectionKind { get; }

        /// <summary>Worksheet that received the imported table.</summary>
        public string SheetName { get; }

        /// <summary>Excel table name requested for the imported range.</summary>
        public string TableName { get; }

        /// <summary>A1 range occupied by the imported Excel table.</summary>
        public string Range { get; }

        /// <summary>Number of imported columns.</summary>
        public int ColumnCount { get; }

        /// <summary>Number of body rows written to Excel.</summary>
        public int RowCount { get; }

        /// <summary>Total body rows detected before any row cap was applied.</summary>
        public int TotalRowCount { get; }

        /// <summary>True when imported rows were truncated by the configured row cap.</summary>
        public bool Truncated { get; }
    }
}
