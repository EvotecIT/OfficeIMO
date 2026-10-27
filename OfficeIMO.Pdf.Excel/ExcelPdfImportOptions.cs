using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf.Excel {
    /// <summary>
    /// Options for extracting logical PDF tables into an Excel workbook.
    /// </summary>
public sealed class ExcelPdfImportOptions {
    /// <summary>Blocks conversion when fidelity loss is reported.</summary>
    public bool FailOnLoss { get; set; }

        /// <summary>Cancellation observed at table, row, and column-analysis boundaries.</summary>
        internal CancellationToken CancellationToken { get; set; }

        /// <summary>
        /// Canonical semantic-read settings used when importing an opened <see cref="OfficeIMO.Pdf.PdfDocument"/>.
        /// Null uses <see cref="OfficeIMO.Pdf.PdfLoadOptions.Default"/>. This setting is ignored when the source is already
        /// an <see cref="OfficeIMO.Pdf.PdfDocumentReadResult"/>.
        /// </summary>
        public OfficeIMO.Pdf.PdfLoadOptions? ReadOptions { get; set; }

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
        /// With <see cref="PreserveColumnProportions"/>, the fitted total supplies the width budget
        /// before detected proportions determine individual column widths.
        /// </summary>
        public bool AutoFitColumns { get; set; } = true;

        /// <summary>
        /// When true, redistributes the worksheet column-width budget using the detected PDF column proportions.
        /// Default false retains content-based auto-fit or worksheet default widths.
        /// </summary>
        /// <remarks>
        /// <see cref="AutoFitColumns"/> selects a content-fitted total or, when false, the worksheet default-width total.
        /// Detected proportions take precedence over individual content-fit widths and may leave some text unfitted.
        /// If any mapped width would exceed Excel's 255-character limit, one common scale reduces the total to retain the proportions.
        /// Canonical visual column strips supply the proportions when available; otherwise native column ranges are used.
        /// Missing, mismatched, non-positive or non-finite column geometry retains the existing widths.
        /// Quarter-turn tables use their visual progression axis. Widths are relative Excel character units;
        /// physical PDF dimensions, authored ruling boundaries and exact spreadsheet printer geometry are not guaranteed.
        /// ODS conversion uses the same widths with the spreadsheet owner's approximate unit conversion.
        /// </remarks>
        public bool PreserveColumnProportions { get; set; }

        /// <summary>
        /// When true, detected numeric PDF table columns are written as numeric Excel cells when every non-empty value can be parsed.
        /// </summary>
        public bool ConvertNumericColumns { get; set; } = true;

        /// <summary>
        /// When true, columns containing only invariant true/false values are written as boolean Excel cells.
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
        /// Optional culture used when parsing dates, date-times, and clock times. When null, date conversion accepts only
        /// unambiguous invariant year-first forms and clock times use invariant culture.
        /// </summary>
        public CultureInfo? DateTimeCulture { get; set; }

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

        /// <summary>Use the first row of a table without an established schema as caller-confirmed column headers.
        /// Default false preserves all source rows. This applies to each imported logical table.</summary>
        public bool UseFirstRowAsHeader { get; set; }

        internal ExcelPdfImportOptions Snapshot() {
            var copy = Clone();
            copy.ReadOptions = ReadOptions?.Clone();
            copy.NumericCulture = CultureInfo.ReadOnly((CultureInfo)NumericCulture.Clone());
            copy.DateTimeCulture = DateTimeCulture == null ? null : CultureInfo.ReadOnly((CultureInfo)DateTimeCulture.Clone());
            return copy;
        }

        /// <summary>Copies the import settings so convenience parameters do not change caller-owned options.</summary>
        /// <remarks>Referenced read settings and cultures are shared; scalar import settings are independent.</remarks>
        public ExcelPdfImportOptions Clone() =>
            (ExcelPdfImportOptions)MemberwiseClone();
    }
}
