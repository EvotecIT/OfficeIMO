namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Describes one logical PDF table imported into an Excel worksheet.
    /// </summary>
    public sealed class PdfExcelTableImportEntry {
        internal PdfExcelTableImportEntry(
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

    /// <summary>Reports the tables imported while converting a logical PDF to an Excel workbook.</summary>
    public sealed class PdfExcelConversionReport {
        internal PdfExcelConversionReport(IReadOnlyList<PdfExcelTableImportEntry> entries) {
            Entries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
        }

        /// <summary>Gets a snapshot of imported table metadata.</summary>
        public IReadOnlyList<PdfExcelTableImportEntry> Entries { get; }

        /// <summary>Gets whether any source table was truncated by the configured row limit.</summary>
        public bool HasLoss => Entries.Any(static entry => entry.Truncated);

        /// <summary>Throws when at least one source table was truncated.</summary>
        public void RequireNoLoss() {
            if (HasLoss) throw new InvalidOperationException("PDF-to-Excel conversion truncated one or more source tables.");
        }
    }

    /// <summary>Contains an editable Excel document and the corresponding PDF table conversion report.</summary>
    public sealed class PdfExcelConversionResult {
        internal PdfExcelConversionResult(ExcelDocument value, PdfExcelConversionReport report) {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the generated editable Excel document. The caller owns and disposes it.</summary>
        public ExcelDocument Value { get; }

        /// <summary>Gets the immutable conversion report.</summary>
        public PdfExcelConversionReport Report { get; }

        /// <summary>Gets whether the conversion truncated source content.</summary>
        public bool HasLoss => Report.HasLoss;

        /// <summary>Returns the generated editable Excel document.</summary>
        public ExcelDocument RequireValue() => Value;

        /// <summary>Returns the generated editable workbook only when no truncation was reported.</summary>
        public ExcelDocument RequireNoLoss() {
            Report.RequireNoLoss();
            return Value;
        }
    }
}
