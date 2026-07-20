using PdfCore = OfficeIMO.Pdf;
using DrawingCore = OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Options controlling first-party OfficeIMO Excel-to-PDF export.
    /// </summary>
    public sealed class ExcelPdfSaveOptions {
        private int _headerRowCount = 1;
        private int? _maxRowsPerSheet;
        private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault();

        /// <summary>
        /// Warnings populated when workbook content cannot be mapped faithfully.
        /// The collection is cleared at the start of each export.
        /// </summary>
        internal List<ExcelPdfExportWarning> Warnings { get; private set; } = new List<ExcelPdfExportWarning>();

        /// <summary>
        /// Shared conversion report populated alongside <see cref="Warnings"/> for wrapper-friendly diagnostics.
        /// The report is cleared at the start of each export.
        /// </summary>
        internal PdfCore.PdfConversionReport Report { get; private set; } = new PdfCore.PdfConversionReport();

        /// <summary>
        /// PDF creation options passed to the first-party PDF engine. The options are cloned before export.
        /// </summary>
        public PdfCore.PdfOptions? PdfOptions { get; set; }

        /// <summary>
        /// Optional workbook default font family used by the first-party PDF engine.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>Host-resource policy. Defaults to balanced conversion: system fonts and bounded in-source resources are allowed, while local and remote reads are denied.</summary>
        public PdfCore.PdfResourcePolicy ResourcePolicy {
            get => _resourcePolicy;
            set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Built-in generated-text fallback groups applied by the Excel PDF converter.
        /// </summary>
        public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

        /// <summary>
        /// Optional first-party page size in PDF points.
        /// </summary>
        public PdfCore.PageSize? PageSize { get; set; }

        /// <summary>
        /// Optional first-party page margins in PDF points.
        /// </summary>
        public PdfCore.PageMargins? Margins { get; set; }

        /// <summary>
        /// Controls whether sheets preserve worksheet coordinates or are reflowed as report tables.
        /// Defaults to <see cref="ExcelPdfWorksheetLayoutMode.WorksheetCanvas"/>.
        /// </summary>
        public ExcelPdfWorksheetLayoutMode WorksheetLayout { get; set; } = ExcelPdfWorksheetLayoutMode.WorksheetCanvas;

        /// <summary>
        /// Optional workbook sheet names to export. When null or empty, all workbook sheets are exported in workbook order.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }

        /// <summary>
        /// When true, hidden workbook worksheets are omitted from the default all-sheets export. Explicit SheetNames still export requested sheets. Defaults to true.
        /// </summary>
        public bool RespectWorkbookSheetVisibility { get; set; } = true;

        /// <summary>
        /// When true, worksheet print areas are used instead of the used range when configured. Defaults to true.
        /// </summary>
        public bool UseWorksheetPrintAreas { get; set; } = true;

        /// <summary>
        /// When true, worksheet orientation and margins are applied when explicit PDF options do not replace them. Defaults to true.
        /// </summary>
        public bool UseWorksheetPageSetup { get; set; } = true;

        /// <summary>
        /// When true, worksheet repeated print-title rows are exported as PDF table header rows. Defaults to true.
        /// </summary>
        public bool UseWorksheetPrintTitleRows { get; set; } = true;

        /// <summary>
        /// When true, manual worksheet row page breaks split the exported PDF table across pages. Defaults to true.
        /// </summary>
        public bool UseWorksheetPageBreaks { get; set; } = true;

        /// <summary>
        /// When true, simple worksheet header/footer text zones are exported to PDF page header/footer zones. Defaults to true.
        /// </summary>
        public bool UseWorksheetHeadersAndFooters { get; set; } = true;

        /// <summary>
        /// Optional local date/time provider used when expanding Excel header/footer &amp;D and &amp;T fields.
        /// </summary>
        public Func<DateTime>? HeaderFooterDateTimeProvider { get; set; }

        /// <summary>
        /// When true, supported worksheet header/footer images are exported to PDF header/footer image zones. Defaults to true.
        /// </summary>
        public bool UseWorksheetHeaderFooterImages { get; set; } = true;

        /// <summary>
        /// When true, simple worksheet cell number formats, font emphasis, colors, alignment, and borders are exported to PDF table cells. Defaults to true.
        /// </summary>
        public bool UseWorksheetCellStyles { get; set; } = true;

        /// <summary>
        /// When true, external worksheet cell hyperlinks and internal workbook links to exported sheets are exported as PDF table-cell link annotations. Defaults to true.
        /// </summary>
        public bool UseWorksheetHyperlinks { get; set; } = true;

        /// <summary>
        /// When true, supported worksheet drawing images are exported at worksheet coordinates
        /// in canvas mode or as flow images in compatibility mode. Defaults to true.
        /// </summary>
        public bool UseWorksheetImages { get; set; } = true;

        /// <summary>
        /// When true, supported worksheet charts are exported as first-party PDF drawing snapshots. Defaults to true.
        /// </summary>
        public bool UseWorksheetCharts { get; set; } = true;

        /// <summary>
        /// Optional shared chart style applied to exported worksheet chart snapshots and generated chart legend tables.
        /// </summary>
        public DrawingCore.OfficeChartStyle? ChartStyle { get; set; }

        /// <summary>
        /// Optional shared chart layout applied to exported worksheet chart snapshots.
        /// </summary>
        public DrawingCore.OfficeChartLayout? ChartLayout { get; set; }

        /// <summary>
        /// When true, worksheet merged cells are exported as PDF table column and row spans. Defaults to true.
        /// </summary>
        public bool UseWorksheetMergedCells { get; set; } = true;

        /// <summary>
        /// When true, explicit worksheet column widths influence PDF table column proportions. Defaults to true.
        /// </summary>
        public bool UseWorksheetColumnWidths { get; set; } = true;

        /// <summary>
        /// When true, explicit worksheet row heights influence PDF table row heights. Defaults to true.
        /// </summary>
        public bool UseWorksheetRowHeights { get; set; } = true;

        /// <summary>
        /// When true, hidden worksheet rows and columns are omitted from the exported PDF table. Defaults to true.
        /// </summary>
        public bool RespectWorksheetHiddenRowsAndColumns { get; set; } = true;

        /// <summary>
        /// Determines whether exported sheets start with a generated worksheet-name heading. Defaults to false.
        /// </summary>
        public bool IncludeSheetHeadings { get; set; }

        /// <summary>
        /// Number of leading worksheet rows styled as table headers. Defaults to one row.
        /// </summary>
        public int HeaderRowCount {
            get => _headerRowCount;
            set {
                if (value < 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Header row count cannot be negative.");
                }

                _headerRowCount = value;
            }
        }

        /// <summary>
        /// Optional maximum number of used-range rows exported from each sheet.
        /// </summary>
        public int? MaxRowsPerSheet {
            get => _maxRowsPerSheet;
            set {
                if (value.HasValue && value.Value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Maximum exported row count must be positive.");
                }

                _maxRowsPerSheet = value;
            }
        }

        /// <summary>
        /// Text used for empty worksheet cells in the exported PDF table. Defaults to an empty string.
        /// </summary>
        public string EmptyCellText { get; set; } = string.Empty;

        /// <summary>
        /// Applies a high-level export profile by setting the Excel PDF options that correspond to that profile.
        /// </summary>
        public ExcelPdfSaveOptions UseProfile(PdfCore.PdfExportProfile profile) {
            switch (profile) {
                case PdfCore.PdfExportProfile.Faithful:
                    WorksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas;
                    RespectWorkbookSheetVisibility = true;
                    UseWorksheetPrintAreas = true;
                    UseWorksheetPageSetup = true;
                    UseWorksheetPrintTitleRows = true;
                    UseWorksheetPageBreaks = true;
                    UseWorksheetHeadersAndFooters = true;
                    UseWorksheetHeaderFooterImages = true;
                    UseWorksheetCellStyles = true;
                    UseWorksheetHyperlinks = true;
                    UseWorksheetImages = true;
                    UseWorksheetCharts = true;
                    UseWorksheetMergedCells = true;
                    UseWorksheetColumnWidths = true;
                    UseWorksheetRowHeights = true;
                    RespectWorksheetHiddenRowsAndColumns = true;
                    IncludeSheetHeadings = false;
                    break;
                case PdfCore.PdfExportProfile.Lightweight:
                    WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable;
                    RespectWorkbookSheetVisibility = true;
                    UseWorksheetPrintAreas = true;
                    UseWorksheetPageSetup = true;
                    UseWorksheetPrintTitleRows = true;
                    UseWorksheetPageBreaks = true;
                    UseWorksheetHeadersAndFooters = true;
                    UseWorksheetHeaderFooterImages = false;
                    UseWorksheetCellStyles = true;
                    UseWorksheetHyperlinks = false;
                    UseWorksheetImages = false;
                    UseWorksheetCharts = false;
                    UseWorksheetMergedCells = true;
                    UseWorksheetColumnWidths = true;
                    UseWorksheetRowHeights = true;
                    RespectWorksheetHiddenRowsAndColumns = true;
                    IncludeSheetHeadings = false;
                    break;
                case PdfCore.PdfExportProfile.PrintReady:
                    WorksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas;
                    RespectWorkbookSheetVisibility = true;
                    UseWorksheetPrintAreas = true;
                    UseWorksheetPageSetup = true;
                    UseWorksheetPrintTitleRows = true;
                    UseWorksheetPageBreaks = true;
                    UseWorksheetHeadersAndFooters = true;
                    UseWorksheetHeaderFooterImages = true;
                    UseWorksheetCellStyles = true;
                    UseWorksheetHyperlinks = true;
                    UseWorksheetImages = true;
                    UseWorksheetCharts = true;
                    UseWorksheetMergedCells = true;
                    UseWorksheetColumnWidths = true;
                    UseWorksheetRowHeights = true;
                    RespectWorksheetHiddenRowsAndColumns = true;
                    IncludeSheetHeadings = false;
                    break;
                case PdfCore.PdfExportProfile.TextOnly:
                    WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable;
                    RespectWorkbookSheetVisibility = true;
                    UseWorksheetPrintAreas = true;
                    UseWorksheetPageSetup = true;
                    UseWorksheetPrintTitleRows = true;
                    UseWorksheetPageBreaks = true;
                    UseWorksheetHeadersAndFooters = true;
                    UseWorksheetHeaderFooterImages = false;
                    UseWorksheetCellStyles = false;
                    UseWorksheetHyperlinks = false;
                    UseWorksheetImages = false;
                    UseWorksheetCharts = false;
                    UseWorksheetMergedCells = true;
                    UseWorksheetColumnWidths = true;
                    UseWorksheetRowHeights = true;
                    RespectWorksheetHiddenRowsAndColumns = true;
                    IncludeSheetHeadings = false;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported PDF export profile.");
            }

            return this;
        }

        internal ExcelPdfSaveOptions CloneForConversion() {
            var clone = (ExcelPdfSaveOptions)MemberwiseClone();
            clone.SheetNames = SheetNames?.ToArray();
            clone.ResourcePolicy = ResourcePolicy.Clone();
            clone.Warnings = new List<ExcelPdfExportWarning>();
            clone.Report = new PdfCore.PdfConversionReport();
            return clone;
        }
    }
}
