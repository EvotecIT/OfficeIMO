using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for dependency-free Excel range, worksheet, and workbook image export.
    /// </summary>
    public class ExcelImageExportOptions : OfficeImageExportOptions {
        /// <summary>Default maximum number of worksheet cells materialized for one visual snapshot.</summary>
        public const int DefaultMaximumRenderedCells = 100_000;

        /// <summary>Default maximum aggregate source-image bytes read for one visual snapshot.</summary>
        public const long DefaultMaximumTotalSourceImageBytes = 256L * 1024L * 1024L;

        /// <summary>
        /// Gridline color used when <see cref="ShowGridlines"/> is enabled.
        /// </summary>
        public OfficeColor GridlineColor { get; set; } = OfficeColor.FromRgb(217, 217, 217);

        /// <summary>
        /// Whether default worksheet gridlines are rendered.
        /// </summary>
        public bool ShowGridlines { get; set; } = true;

        /// <summary>
        /// Whether hidden rows and columns should be included.
        /// </summary>
        public bool IncludeHidden { get; set; }

        /// <summary>
        /// Whether worksheet images should be included.
        /// </summary>
        public bool IncludeImages { get; set; } = true;

        /// <summary>
        /// Whether worksheet charts should be included when supported.
        /// </summary>
        public bool IncludeCharts { get; set; } = true;

        /// <summary>
        /// Whether supported worksheet drawing objects, such as simple shapes and text boxes, should be included.
        /// </summary>
        public bool IncludeDrawingObjects { get; set; } = true;

        /// <summary>
        /// Whether supported conditional formatting visuals should be included.
        /// </summary>
        public bool IncludeConditionalFormatting { get; set; } = true;

        /// <summary>
        /// Date used to evaluate relative conditional-formatting time-period rules such as Today, Yesterday, Last7Days, and ThisMonth.
        /// When omitted, image export uses the current local date captured during option normalization.
        /// </summary>
        public DateTime? ConditionalFormattingDate { get; set; }

        /// <summary>
        /// Whether cells with hyperlink metadata should get a default blue underline hint when their style does not already make the link visible.
        /// </summary>
        public bool ShowHyperlinkHints { get; set; } = true;

        /// <summary>
        /// Whether visible cell comment and threaded-comment bodies should be rendered as dependency-free callouts.
        /// </summary>
        public bool ShowCommentBodies { get; set; }

        /// <summary>
        /// Default column width in pixels when a column has no explicit width.
        /// </summary>
        public double DefaultColumnWidthPixels { get; set; } = 64D;

        /// <summary>
        /// Default row height in pixels when a row has no explicit height.
        /// </summary>
        public double DefaultRowHeightPixels { get; set; } = 20D;

        /// <summary>
        /// Maximum number of worksheet cells materialized for one visual snapshot.
        /// This limit is applied before rows, columns, and cells are expanded from drawing-driven ranges.
        /// </summary>
        public int MaximumRenderedCells { get; set; } = DefaultMaximumRenderedCells;

        /// <summary>
        /// Maximum aggregate bytes read from embedded worksheet images for one visual snapshot.
        /// This is independent from <see cref="OfficeImageExportOptions.MaximumTotalEncodedBytes"/>,
        /// which limits encoded output bytes.
        /// </summary>
        public long MaximumTotalSourceImageBytes { get; set; } = DefaultMaximumTotalSourceImageBytes;

        /// <summary>Creates an independent options snapshot.</summary>
        public ExcelImageExportOptions Clone() => CopyExcelOptionsTo(new ExcelImageExportOptions());

        internal T CopyExcelOptionsTo<T>(T target) where T : ExcelImageExportOptions {
            CopyImageExportOptionsTo(target);
            target.GridlineColor = GridlineColor;
            target.ShowGridlines = ShowGridlines;
            target.IncludeHidden = IncludeHidden;
            target.IncludeImages = IncludeImages;
            target.IncludeCharts = IncludeCharts;
            target.IncludeDrawingObjects = IncludeDrawingObjects;
            target.IncludeConditionalFormatting = IncludeConditionalFormatting;
            target.ConditionalFormattingDate = ConditionalFormattingDate;
            target.ShowHyperlinkHints = ShowHyperlinkHints;
            target.ShowCommentBodies = ShowCommentBodies;
            target.DefaultColumnWidthPixels = DefaultColumnWidthPixels;
            target.DefaultRowHeightPixels = DefaultRowHeightPixels;
            target.MaximumRenderedCells = MaximumRenderedCells;
            target.MaximumTotalSourceImageBytes = MaximumTotalSourceImageBytes;
            return target;
        }

        internal void Validate() {
            ValidateImageExportOptions();
            if (MaximumRenderedCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaximumRenderedCells), "Maximum rendered cells must be positive.");
            }
            if (MaximumTotalSourceImageBytes <= 0L) {
                throw new ArgumentOutOfRangeException(nameof(MaximumTotalSourceImageBytes), "Maximum total source-image bytes must be positive.");
            }
        }
    }

    /// <summary>
    /// Options for worksheet image export.
    /// </summary>
    public sealed class ExcelWorksheetImageExportOptions : ExcelImageExportOptions {
        /// <summary>
        /// Optional explicit range. When omitted, the worksheet used range is exported.
        /// </summary>
        public string? Range { get; set; }

        /// <summary>
        /// Timestamp used for dynamic Excel header/footer date and time fields such as <c>&amp;D</c>, <c>&amp;T</c>, <c>&amp;[Date]</c>, and <c>&amp;[Time]</c>.
        /// When omitted, image export uses the current local time captured at the start of the worksheet export.
        /// </summary>
        public DateTime? HeaderFooterDateTime { get; set; }

        /// <summary>
        /// When true and <see cref="Range"/> is omitted, uses the worksheet print area when configured.
        /// </summary>
        public bool UsePrintArea { get; set; }

        /// <summary>
        /// When true, <see cref="ExcelSheet.ExportImages(OfficeImageExportFormat, ExcelWorksheetImageExportOptions?)"/> splits output at manual row and column page breaks.
        /// Single-image worksheet export keeps one result and emits a diagnostic instead of splitting.
        /// </summary>
        public bool SplitByManualPageBreaks { get; set; }

        /// <summary>Creates an independent worksheet options snapshot.</summary>
        public ExcelWorksheetImageExportOptions CloneWorksheet() {
            ExcelWorksheetImageExportOptions clone = CopyExcelOptionsTo(new ExcelWorksheetImageExportOptions());
            clone.Range = Range;
            clone.HeaderFooterDateTime = HeaderFooterDateTime;
            clone.UsePrintArea = UsePrintArea;
            clone.SplitByManualPageBreaks = SplitByManualPageBreaks;
            return clone;
        }
    }

    /// <summary>
    /// Options for workbook image export.
    /// </summary>
    public sealed class ExcelWorkbookImageExportOptions : ExcelImageExportOptions {
        /// <summary>
        /// Optional list of worksheet names to export. When omitted, visible worksheets are exported.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>
        /// When true and <see cref="SheetNames"/> is omitted, workbook image export includes hidden and very hidden worksheets.
        /// Explicitly named worksheets are exported regardless of visibility.
        /// </summary>
        public bool IncludeHiddenSheets { get; set; }

        /// <summary>
        /// Timestamp used for dynamic Excel header/footer date and time fields in worksheet image exports.
        /// When omitted, workbook image export uses the current local time captured at the start of the workbook export.
        /// </summary>
        public DateTime? HeaderFooterDateTime { get; set; }

        /// <summary>
        /// When true, each worksheet image export uses that worksheet's print area when configured.
        /// </summary>
        public bool UseWorksheetPrintAreas { get; set; }

        /// <summary>
        /// When true, workbook image export asks each worksheet to split output at manual row and column page breaks.
        /// </summary>
        public bool SplitWorksheetsByManualPageBreaks { get; set; }

        /// <summary>Creates an independent workbook options snapshot.</summary>
        public ExcelWorkbookImageExportOptions CloneWorkbook() {
            ExcelWorkbookImageExportOptions clone = CopyExcelOptionsTo(new ExcelWorkbookImageExportOptions());
            clone.SheetNames = SheetNames?.ToArray();
            clone.IncludeHiddenSheets = IncludeHiddenSheets;
            clone.HeaderFooterDateTime = HeaderFooterDateTime;
            clone.UseWorksheetPrintAreas = UseWorksheetPrintAreas;
            clone.SplitWorksheetsByManualPageBreaks = SplitWorksheetsByManualPageBreaks;
            return clone;
        }
    }
}
