namespace OfficeIMO.Excel.Pdf {
    /// <summary>Typed value kind selected for an imported Excel table column.</summary>
    public enum PdfExcelTableColumnKind {
        /// <summary>Text cells.</summary>
        Text,
        /// <summary>Numeric cells.</summary>
        Number,
        /// <summary>Fractional numeric cells parsed from percentage text.</summary>
        Percentage,
        /// <summary>Boolean cells.</summary>
        Boolean,
        /// <summary>Date or date-time cells.</summary>
        DateTime,
        /// <summary>Time-of-day cells without an invented calendar date.</summary>
        Time
    }

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
            bool truncated,
            IReadOnlyList<int> sourcePageNumbers,
            int sourceTableCount,
            int suppressedRepeatedHeaderRows,
            int additionalHeaderRowCount,
            IReadOnlyList<PdfExcelTableColumnKind> columnKinds) {
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
            SourcePageNumbers = Array.AsReadOnly(sourcePageNumbers.ToArray());
            SourceTableCount = sourceTableCount;
            SuppressedRepeatedHeaderRows = suppressedRepeatedHeaderRows;
            AdditionalHeaderRowCount = additionalHeaderRowCount;
            ColumnKinds = Array.AsReadOnly(columnKinds.ToArray());
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

        /// <summary>One-based PDF page numbers contributing rows to this imported table.</summary>
        public IReadOnlyList<int> SourcePageNumbers { get; }

        /// <summary>Number of detected page-level table segments combined into this imported table.</summary>
        public int SourceTableCount { get; }

        /// <summary>Number of repeated continuation header rows omitted from body data.</summary>
        public int SuppressedRepeatedHeaderRows { get; }

        /// <summary>Number of repeated header rows appended to the primary header labels.</summary>
        public int AdditionalHeaderRowCount { get; }

        /// <summary>Typed value kinds selected for the imported columns.</summary>
        public IReadOnlyList<PdfExcelTableColumnKind> ColumnKinds { get; }
    }

    /// <summary>Reports the detected tables imported from a logical PDF into an Excel workbook.</summary>
    public sealed class PdfExcelTableImportReport : IOfficeConversionReport {
        internal PdfExcelTableImportReport(
            IReadOnlyList<PdfExcelTableImportEntry> entries,
            OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope) {
            Entries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
            SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
        }

        /// <summary>Gets a snapshot of imported table metadata.</summary>
        public IReadOnlyList<PdfExcelTableImportEntry> Entries { get; }

        /// <summary>Gets source-page content that was outside this table-only import.</summary>
        public OfficeIMO.Pdf.PdfTableExtractionScopeReport SourceScope { get; }

        /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
        public bool HasOmittedPageContent => SourceScope.HasOmittedPageContent;

        /// <summary>Gets whether any detected source table was truncated by the configured row limit.</summary>
        public bool HasLoss => Entries.Any(static entry => entry.Truncated);

        /// <summary>Throws when at least one detected source table was truncated.</summary>
        public void RequireNoLoss() {
            if (HasLoss) throw new InvalidOperationException("PDF table import to Excel truncated one or more detected source tables.");
        }
    }

    /// <summary>Contains an editable Excel document and the corresponding PDF table import report.</summary>
    public sealed class PdfExcelTableImportResult : OfficeConversionResult<ExcelDocument, PdfExcelTableImportReport> {
        internal PdfExcelTableImportResult(ExcelDocument value, PdfExcelTableImportReport report) : base(value, report) { }

        /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
        public bool HasOmittedPageContent => Report.HasOmittedPageContent;

    }
}
