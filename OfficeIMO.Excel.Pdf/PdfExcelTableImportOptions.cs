using System.Globalization;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Options for extracting logical PDF tables into an Excel workbook.
    /// </summary>
public sealed class PdfExcelTableImportOptions {
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
        public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium2;

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
        /// When true, columns containing only boolean values such as true/false or yes/no are written as boolean Excel cells.
        /// </summary>
        public bool ConvertBooleanColumns { get; set; } = true;

        /// <summary>
        /// When true, unambiguous date or date-time columns are written as date/time Excel cells.
        /// </summary>
        public bool ConvertDateTimeColumns { get; set; } = true;

        /// <summary>
        /// When true, columns containing only percentage values are written as fractional numeric Excel cells.
        /// </summary>
        public bool ConvertPercentageColumns { get; set; } = true;

        /// <summary>
        /// Culture used when parsing detected numeric PDF table values before writing typed Excel cells.
        /// </summary>
        public CultureInfo NumericCulture { get; set; } = CultureInfo.InvariantCulture;

        /// <summary>
        /// When true, adjacent page-edge table segments with compatible geometry and schema are imported as one logical table.
        /// </summary>
        public bool MergePageContinuations { get; set; } = true;

        /// <summary>
        /// When true, identical header-like body prefixes on every merged segment are treated as additional repeated header rows and suppressed.
        /// Keep false when repeated ordinary data rows must be preserved without additional header evidence.
        /// </summary>
        public bool SuppressRepeatedBodyHeaderRows { get; set; }

        /// <summary>
        /// Maximum adjacent PDF table segments that may be combined into one logical table.
        /// </summary>
        public int MaximumContinuationSegments { get; set; } = 64;

        /// <summary>
        /// Maximum per-column geometry difference, in PDF points, allowed when recognizing a page continuation.
        /// </summary>
        public double ContinuationGeometryTolerancePoints { get; set; } = 4D;

        /// <summary>
        /// Worksheet name used when no tables are detected, keeping the produced workbook valid.
        /// </summary>
        public string EmptyWorkbookSheetName { get; set; } = "PDF Tables";
    }
}
