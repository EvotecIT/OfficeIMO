using PdfCore = OfficeIMO.Pdf;
using System.Globalization;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Options for extracting logical PDF tables into an Excel workbook.
    /// </summary>
    public sealed class PdfExcelTableImportOptions {
        /// <summary>
        /// PDF text layout options used when a path, stream, or byte array is loaded directly.
        /// </summary>
        public PdfCore.PdfTextLayoutOptions? LayoutOptions { get; set; }

        /// <summary>
        /// Optional inclusive one-based source page ranges used by direct PDF loading overloads.
        /// </summary>
        public IReadOnlyList<PdfCore.PdfPageRange>? PageRanges { get; set; }

        /// <summary>
        /// Maximum body rows to import per detected table. Values less than or equal to zero import all rows.
        /// </summary>
        public int MaxRows { get; set; }

        /// <summary>
        /// Worksheet name prefix used before the source page and table coordinates.
        /// </summary>
        public string SheetNamePrefix { get; set; } = "PDF";

        /// <summary>
        /// Excel table name prefix used before the source page and table coordinates.
        /// </summary>
        public string TableNamePrefix { get; set; } = "PdfTable";

        /// <summary>
        /// Excel table style applied to imported tables.
        /// </summary>
        public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium2;

        /// <summary>
        /// When true, created Excel tables include a table-scoped AutoFilter.
        /// </summary>
        public bool IncludeAutoFilter { get; set; } = true;

        /// <summary>
        /// When true, worksheet columns are auto-fitted after the table is inserted.
        /// </summary>
        public bool AutoFitColumns { get; set; } = true;

        /// <summary>
        /// When true, detected numeric PDF table columns are written as numeric Excel cells when every non-empty value can be parsed.
        /// </summary>
        public bool ConvertNumericColumns { get; set; } = true;

        /// <summary>
        /// Culture used when parsing detected numeric PDF table values before writing typed Excel cells.
        /// </summary>
        public CultureInfo NumericCulture { get; set; } = CultureInfo.InvariantCulture;

        /// <summary>
        /// Worksheet name used when no tables are detected, keeping the produced workbook valid.
        /// </summary>
        public string EmptyWorkbookSheetName { get; set; } = "PDF Tables";
    }
}
