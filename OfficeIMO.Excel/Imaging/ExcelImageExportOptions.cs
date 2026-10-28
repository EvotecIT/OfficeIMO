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

        /// <summary>Default maximum number of images produced from manual page-break splitting.</summary>
        public const int DefaultMaximumPageBreakImages = 1_000;

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
        /// Whether workbook directory paths referenced by header/footer <c>&amp;Z</c> and <c>&amp;[Path]</c>
        /// fields may be included in rendered images. The default is <see langword="false"/> to avoid
        /// exposing local or server filesystem paths in exported artifacts.
        /// </summary>
        public bool IncludeWorkbookPathInHeaderFooter { get; set; }

        /// <summary>
        /// Fallback physical column width used by image layout when no explicit width is stored.
        /// </summary>
        public OfficeLength DefaultColumnWidth { get; set; } = OfficeLength.Pixels(64);

        /// <summary>
        /// Fallback physical row height used by image layout when no explicit height is stored.
        /// </summary>
        public OfficeLength DefaultRowHeight { get; set; } = OfficeLength.Pixels(20);

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

        /// <summary>
        /// Maximum aggregate number of image results that multi-area print ranges and manual row or column page breaks may produce for one worksheet export.
        /// </summary>
        public int MaximumPageBreakImages { get; set; } = DefaultMaximumPageBreakImages;

        /// <summary>Explicit A1 range, or null to render each selected worksheet's used range.</summary>
        public string? Range { get; set; }

        /// <summary>Timestamp for dynamic header/footer fields; captured once per operation when omitted.</summary>
        public DateTime? HeaderFooterDateTime { get; set; }

        /// <summary>Uses configured print-area segments when no explicit range is supplied.</summary>
        public bool UsePrintArea { get; set; }

        /// <summary>Splits worksheet batches at manual row and column page breaks. Single-image rendering reports unsplit content.</summary>
        public bool SplitByManualPageBreaks { get; set; }

        /// <summary>Worksheet names to export. Explicitly selected sheets are included regardless of visibility.</summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>Includes hidden and very hidden worksheets when no explicit sheet selection is supplied.</summary>
        public bool IncludeHiddenSheets { get; set; }

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
            target.IncludeWorkbookPathInHeaderFooter = IncludeWorkbookPathInHeaderFooter;
            target.DefaultColumnWidth = DefaultColumnWidth;
            target.DefaultRowHeight = DefaultRowHeight;
            target.MaximumRenderedCells = MaximumRenderedCells;
            target.MaximumTotalSourceImageBytes = MaximumTotalSourceImageBytes;
            target.MaximumPageBreakImages = MaximumPageBreakImages;
            target.Range = Range;
            target.HeaderFooterDateTime = HeaderFooterDateTime;
            target.UsePrintArea = UsePrintArea;
            target.SplitByManualPageBreaks = SplitByManualPageBreaks;
            target.SheetNames = SheetNames?.ToArray();
            target.IncludeHiddenSheets = IncludeHiddenSheets;
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
            if (MaximumPageBreakImages <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaximumPageBreakImages), "Maximum page-break images must be positive.");
            }
        }
    }

}
