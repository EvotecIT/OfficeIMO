namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Describes one logical PDF table imported into a Word document.
    /// </summary>
    public sealed class PdfWordTableImportResult {
        internal PdfWordTableImportResult(
            int pageIndex,
            int pageNumber,
            int tableIndex,
            string detectionKind,
            int columnCount,
            int rowCount,
            int totalRowCount,
            bool truncated,
            bool headerRowIncluded) {
            PageIndex = pageIndex;
            PageNumber = pageNumber;
            TableIndex = tableIndex;
            DetectionKind = detectionKind ?? string.Empty;
            ColumnCount = columnCount;
            RowCount = rowCount;
            TotalRowCount = totalRowCount;
            Truncated = truncated;
            HeaderRowIncluded = headerRowIncluded;
        }

        /// <summary>Zero-based page index within the selected logical page collection.</summary>
        public int PageIndex { get; }

        /// <summary>One-based source page number from the PDF document.</summary>
        public int PageNumber { get; }

        /// <summary>Zero-based table index within the source logical PDF page.</summary>
        public int TableIndex { get; }

        /// <summary>Detection heuristic that produced the imported table.</summary>
        public string DetectionKind { get; }

        /// <summary>Number of imported columns.</summary>
        public int ColumnCount { get; }

        /// <summary>Number of body rows written to Word.</summary>
        public int RowCount { get; }

        /// <summary>Total body rows detected before any row cap was applied.</summary>
        public int TotalRowCount { get; }

        /// <summary>True when imported rows were truncated by the configured row cap.</summary>
        public bool Truncated { get; }

        /// <summary>True when a column-header row was written above the imported body rows.</summary>
        public bool HeaderRowIncluded { get; }
    }
}
