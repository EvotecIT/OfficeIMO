using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for dependency-free Excel range, worksheet, and workbook image export.
    /// </summary>
    public class ExcelImageExportOptions {
        /// <summary>
        /// Output scale multiplier. A value of 2 creates a 2x PNG/SVG surface.
        /// </summary>
        public double Scale { get; set; } = 1D;

        /// <summary>
        /// Background color used behind the rendered worksheet range.
        /// </summary>
        public OfficeColor BackgroundColor { get; set; } = OfficeColor.White;

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
        /// Whether worksheet images should be included when supported.
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

        internal ExcelImageExportOptions Clone() => new ExcelImageExportOptions {
            Scale = Scale,
            BackgroundColor = BackgroundColor,
            GridlineColor = GridlineColor,
            ShowGridlines = ShowGridlines,
            IncludeHidden = IncludeHidden,
            IncludeImages = IncludeImages,
            IncludeCharts = IncludeCharts,
            IncludeDrawingObjects = IncludeDrawingObjects,
            IncludeConditionalFormatting = IncludeConditionalFormatting,
            ShowHyperlinkHints = ShowHyperlinkHints,
            ShowCommentBodies = ShowCommentBodies,
            DefaultColumnWidthPixels = DefaultColumnWidthPixels,
            DefaultRowHeightPixels = DefaultRowHeightPixels
        };
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
        /// When true and <see cref="Range"/> is omitted, uses the worksheet print area when configured.
        /// </summary>
        public bool UsePrintArea { get; set; }

        /// <summary>
        /// When true, <see cref="ExcelSheet.ExportImages(OfficeImageExportFormat, ExcelWorksheetImageExportOptions?)"/> splits output at manual row and column page breaks.
        /// Single-image worksheet export keeps one result and emits a diagnostic instead of splitting.
        /// </summary>
        public bool SplitByManualPageBreaks { get; set; }
    }

    /// <summary>
    /// Options for workbook image export.
    /// </summary>
    public sealed class ExcelWorkbookImageExportOptions : ExcelImageExportOptions {
        /// <summary>
        /// Optional list of worksheet names to export. When omitted, all worksheets are exported.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>
        /// When true, each worksheet image export uses that worksheet's print area when configured.
        /// </summary>
        public bool UseWorksheetPrintAreas { get; set; }

        /// <summary>
        /// When true, workbook image export asks each worksheet to split output at manual row and column page breaks.
        /// </summary>
        public bool SplitWorksheetsByManualPageBreaks { get; set; }
    }
}
