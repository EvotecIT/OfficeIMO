namespace OfficeIMO.Pdf;

/// <summary>
/// Describes visual and layout options for table rendering.
/// Attach an instance to a table block or use the presets in <see cref="TableStyles"/>.
/// </summary>
public class PdfTableStyle {
    private PdfAlign _captionAlign = PdfAlign.Left;
    private System.Collections.Generic.List<PdfColumnAlign>? _alignments;
    private System.Collections.Generic.List<PdfCellVerticalAlign>? _verticalAlignments;
    private System.Collections.Generic.List<double?>? _columnWidthPoints;
    private System.Collections.Generic.List<double?>? _columnMinWidthPoints;
    private System.Collections.Generic.List<double?>? _columnMaxWidthPoints;
    private System.Collections.Generic.List<double>? _columnWidthWeights;
    private System.Collections.Generic.List<PdfColor?>? _bodyColumnFills;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor>? _cellFills;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellDataBar>? _cellDataBars;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellIcon>? _cellIcons;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellBorder>? _cellBorders;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>? _cellPaddings;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>? _cellAlignments;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>? _cellVerticalAlignments;
    private System.Collections.Generic.List<double?>? _rowMinHeights;
    private System.Collections.Generic.List<double?>? _fixedRowHeights;
    private System.Collections.Generic.List<bool?>? _rowAllowBreakAcrossPages;
    private double _borderWidth = 0.5;
    private double _rowSeparatorWidth;
    private double _headerSeparatorWidth;
    private double _footerSeparatorWidth;
    private int _headerRowCount = 1;
    private int _footerRowCount;
    private int? _repeatHeaderRowCount;
    private int _minimumBodyRowsOnLastPage = 2;
    private double _cellPaddingX = 4;
    private double _cellPaddingY = 2;
    private double? _cellPaddingLeft;
    private double? _cellPaddingRight;
    private double? _cellPaddingTop;
    private double? _cellPaddingBottom;
    private double _cellSpacing;
    private double _minRowHeight;
    private double _spacingBefore;
    private double _pageContinuationSpacingBefore;
    private double? _captionFontSize;
    private double _captionSpacingAfter = 4;
    private double _spacingAfter;
    private double _rowBaselineOffset;
    private double? _preferredWidth;
    private double? _maxWidth;
    private double _leftIndent;
    private double? _fontSize;
    private double? _lineHeight;
    private double? _headerFontSize;
    private double? _footerFontSize;
    private double? _minimumShrinkFontSize;

    /// <summary>Color of the table borders and cell grid lines. Set to null to hide borders.</summary>
    public PdfColor? BorderColor { get; set; } = new PdfColor(0.8, 0.8, 0.8);
    /// <summary>Stroke width, in points, for table borders and cell grid lines.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(BorderWidth), "Table border width must be a non-negative finite value.");
            _borderWidth = value;
        }
    }
    /// <summary>Background fill color for the header row. Set to null for no fill.</summary>
    public PdfColor? HeaderFill { get; set; } = new PdfColor(0.95, 0.95, 0.95);
    /// <summary>Background fill color for footer rows. Set to null for no fill.</summary>
    public PdfColor? FooterFill { get; set; } = new PdfColor(0.95, 0.95, 0.95);
    /// <summary>Optional alternating row fill color (applied to every other body row). Set to null to disable.</summary>
    public PdfColor? RowStripeFill { get; set; } = new PdfColor(0.98, 0.98, 0.98);
    /// <summary>Optional horizontal separator color drawn at the bottom of each table row. Set to null or zero width to disable.</summary>
    public PdfColor? RowSeparatorColor { get; set; }
    /// <summary>Stroke width, in points, for row separators.</summary>
    public double RowSeparatorWidth {
        get => _rowSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(RowSeparatorWidth), "Table row separator width must be a non-negative finite value.");
            _rowSeparatorWidth = value;
        }
    }
    /// <summary>Optional horizontal separator color used below header rows. Falls back to <see cref="RowSeparatorColor"/> when null.</summary>
    public PdfColor? HeaderSeparatorColor { get; set; }
    /// <summary>Optional stroke width, in points, used below header rows. Falls back to <see cref="RowSeparatorWidth"/> when zero.</summary>
    public double HeaderSeparatorWidth {
        get => _headerSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(HeaderSeparatorWidth), "Table header separator width must be a non-negative finite value.");
            _headerSeparatorWidth = value;
        }
    }
    /// <summary>Optional horizontal separator color used above footer rows. Falls back to <see cref="RowSeparatorColor"/> when null.</summary>
    public PdfColor? FooterSeparatorColor { get; set; }
    /// <summary>Optional stroke width, in points, used above footer rows. Falls back to <see cref="RowSeparatorWidth"/> when zero.</summary>
    public double FooterSeparatorWidth {
        get => _footerSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(FooterSeparatorWidth), "Table footer separator width must be a non-negative finite value.");
            _footerSeparatorWidth = value;
        }
    }
    /// <summary>Optional per-column background fills for body rows. Null or missing entries use the row fill.</summary>
    public System.Collections.Generic.List<PdfColor?>? BodyColumnFills {
        get => _bodyColumnFills;
        set => _bodyColumnFills = value == null ? null : new System.Collections.Generic.List<PdfColor?>(value);
    }
    /// <summary>Optional absolute per-cell background fills keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor>? CellFills {
        get => _cellFills;
        set {
            if (value != null) {
                foreach (var cellFill in value) {
                    if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                        throw new System.ArgumentException("Table cell fill coordinates cannot be negative.", nameof(CellFills));
                    }
                }
            }

            _cellFills = value == null ? null : new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor>(value);
        }
    }
    /// <summary>Optional proportional bars drawn inside cells, behind cell text, keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellDataBar>? CellDataBars {
        get => _cellDataBars;
        set {
            if (value == null) {
                _cellDataBars = null;
                return;
            }

            var dataBars = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellDataBar>();
            foreach (var cellDataBar in value) {
                if (cellDataBar.Key.Row < 0 || cellDataBar.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell data bar coordinates cannot be negative.", nameof(CellDataBars));
                }

                if (cellDataBar.Value == null) {
                    throw new System.ArgumentException("Table cell data bars cannot contain null values.", nameof(CellDataBars));
                }

                dataBars[cellDataBar.Key] = cellDataBar.Value.Clone();
            }

            _cellDataBars = dataBars;
        }
    }
    /// <summary>Optional small vector icons drawn inside cells before cell text, keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellIcon>? CellIcons {
        get => _cellIcons;
        set {
            if (value == null) {
                _cellIcons = null;
                return;
            }

            var icons = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellIcon>();
            foreach (var cellIcon in value) {
                if (cellIcon.Key.Row < 0 || cellIcon.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell icon coordinates cannot be negative.", nameof(CellIcons));
                }

                if (cellIcon.Value == null) {
                    throw new System.ArgumentException("Table cell icons cannot contain null values.", nameof(CellIcons));
                }

                icons[cellIcon.Key] = cellIcon.Value.Clone();
            }

            _cellIcons = icons;
        }
    }
    /// <summary>Optional side-specific per-cell border overrides keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellBorder>? CellBorders {
        get => _cellBorders;
        set {
            if (value == null) {
                _cellBorders = null;
                return;
            }

            var borders = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellBorder>();
            foreach (var cellBorder in value) {
                if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell border coordinates cannot be negative.", nameof(CellBorders));
                }

                ValidateCellBorder(cellBorder.Value, nameof(CellBorders));
                borders[cellBorder.Key] = cellBorder.Value.Clone();
            }

            _cellBorders = borders;
        }
    }
    /// <summary>Optional per-cell padding overrides keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>? CellPaddings {
        get => _cellPaddings;
        set {
            if (value == null) {
                _cellPaddings = null;
                return;
            }

            var paddings = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>();
            foreach (var cellPadding in value) {
                if (cellPadding.Key.Row < 0 || cellPadding.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell padding coordinates cannot be negative.", nameof(CellPaddings));
                }

                ValidateCellPadding(cellPadding.Value, nameof(CellPaddings));
                paddings[cellPadding.Key] = cellPadding.Value.Clone();
            }

            _cellPaddings = paddings;
        }
    }
    /// <summary>Optional per-cell horizontal alignment overrides keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>? CellAlignments {
        get => _cellAlignments;
        set {
            if (value == null) {
                _cellAlignments = null;
                return;
            }

            var alignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>();
            foreach (var cellAlignment in value) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell alignment coordinates cannot be negative.", nameof(CellAlignments));
                }

                Guard.TableColumnAlign(cellAlignment.Value, nameof(CellAlignments));
                alignments[cellAlignment.Key] = cellAlignment.Value;
            }

            _cellAlignments = alignments;
        }
    }
    /// <summary>Optional per-cell vertical alignment overrides keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>? CellVerticalAlignments {
        get => _cellVerticalAlignments;
        set {
            if (value == null) {
                _cellVerticalAlignments = null;
                return;
            }

            var alignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>();
            foreach (var cellAlignment in value) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new System.ArgumentException("Table cell vertical alignment coordinates cannot be negative.", nameof(CellVerticalAlignments));
                }

                Guard.TableCellVerticalAlign(cellAlignment.Value, nameof(CellVerticalAlignments));
                alignments[cellAlignment.Key] = cellAlignment.Value;
            }

            _cellVerticalAlignments = alignments;
        }
    }
    /// <summary>Text color for body rows. When null the writer’s default text color is used.</summary>
    public PdfColor? TextColor { get; set; }
    /// <summary>Font size for body cells, in points. When null the document default font size is used.</summary>
    public double? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(FontSize), "Table body font size must be a positive finite value.");
            _fontSize = value;
        }
    }
    /// <summary>Line advance multiplier for wrapped cell text. When null the writer uses the default table line height.</summary>
    public double? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(LineHeight), "Table line height must be a positive finite value.");
            _lineHeight = value;
        }
    }
    /// <summary>Text color for header cells. When null the writer’s default text color is used.</summary>
    public PdfColor? HeaderTextColor { get; set; }
    /// <summary>Font size for header cells, in points. When null the body table font size is used.</summary>
    public double? HeaderFontSize {
        get => _headerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(HeaderFontSize), "Table header font size must be a positive finite value.");
            _headerFontSize = value;
        }
    }
    /// <summary>When true, header cells use the bold variant of the document font family.</summary>
    public bool HeaderBold { get; set; } = true;
    /// <summary>Number of leading rows to render as table headers. Defaults to 1.</summary>
    public int HeaderRowCount {
        get => _headerRowCount;
        set {
            ValidateNonNegativeValue(value, nameof(HeaderRowCount), "Table header row count cannot be negative.");
            _headerRowCount = value;
        }
    }
    /// <summary>Number of leading header rows to repeat on following pages. When null, all configured header rows repeat.</summary>
    public int? RepeatHeaderRowCount {
        get => _repeatHeaderRowCount;
        set {
            if (value.HasValue) {
                ValidateNonNegativeValue(value.Value, nameof(RepeatHeaderRowCount), "Table repeating header row count cannot be negative.");
            }
            _repeatHeaderRowCount = value;
        }
    }
    /// <summary>Number of trailing rows to render as table footers. Defaults to 0.</summary>
    public int FooterRowCount {
        get => _footerRowCount;
        set {
            ValidateNonNegativeValue(value, nameof(FooterRowCount), "Table footer row count cannot be negative.");
            _footerRowCount = value;
        }
    }
    /// <summary>
    /// Minimum number of body rows kept with trailing footer rows on the table's final page when the group fits on one page.
    /// Set to zero to disable this pagination constraint. Defaults to two.
    /// </summary>
    public int MinimumBodyRowsOnLastPage {
        get => _minimumBodyRowsOnLastPage;
        set {
            ValidateNonNegativeValue(value, nameof(MinimumBodyRowsOnLastPage), "Table minimum final-page body row count cannot be negative.");
            _minimumBodyRowsOnLastPage = value;
        }
    }
    /// <summary>Text color for footer cells. When null the writer’s default text color is used.</summary>
    public PdfColor? FooterTextColor { get; set; }
    /// <summary>Font size for footer cells, in points. When null the body table font size is used.</summary>
    public double? FooterFontSize {
        get => _footerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(FooterFontSize), "Table footer font size must be a positive finite value.");
            _footerFontSize = value;
        }
    }
    /// <summary>When true, footer cells use the bold variant of the document font family.</summary>
    public bool FooterBold { get; set; } = true;
    /// <summary>Horizontal padding inside each cell, in points.</summary>
    public double CellPaddingX {
        get => _cellPaddingX;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(CellPaddingX), "Table horizontal cell padding must be a non-negative finite value.");
            _cellPaddingX = value;
        }
    }
    /// <summary>Vertical padding inside each cell, in points.</summary>
    public double CellPaddingY {
        get => _cellPaddingY;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(CellPaddingY), "Table vertical cell padding must be a non-negative finite value.");
            _cellPaddingY = value;
        }
    }
    /// <summary>Optional left padding override inside each cell, in points. When null <see cref="CellPaddingX"/> is used.</summary>
    public double? CellPaddingLeft {
        get => _cellPaddingLeft;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(CellPaddingLeft), "Table left cell padding must be a non-negative finite value.");
            _cellPaddingLeft = value;
        }
    }
    /// <summary>Optional right padding override inside each cell, in points. When null <see cref="CellPaddingX"/> is used.</summary>
    public double? CellPaddingRight {
        get => _cellPaddingRight;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(CellPaddingRight), "Table right cell padding must be a non-negative finite value.");
            _cellPaddingRight = value;
        }
    }
    /// <summary>Optional top padding override inside each cell, in points. When null <see cref="CellPaddingY"/> is used.</summary>
    public double? CellPaddingTop {
        get => _cellPaddingTop;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(CellPaddingTop), "Table top cell padding must be a non-negative finite value.");
            _cellPaddingTop = value;
        }
    }
    /// <summary>Optional bottom padding override inside each cell, in points. When null <see cref="CellPaddingY"/> is used.</summary>
    public double? CellPaddingBottom {
        get => _cellPaddingBottom;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(CellPaddingBottom), "Table bottom cell padding must be a non-negative finite value.");
            _cellPaddingBottom = value;
        }
    }
    /// <summary>Optional spacing between adjacent table cells, in points.</summary>
    public double CellSpacing {
        get => _cellSpacing;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(CellSpacing), "Table cell spacing must be a non-negative finite value.");
            _cellSpacing = value;
        }
    }
    /// <summary>Optional minimum row height in points. Set to 0 to size rows from wrapped content.</summary>
    public double MinRowHeight {
        get => _minRowHeight;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(MinRowHeight), "Table minimum row height must be a non-negative finite value.");
            _minRowHeight = value;
        }
    }
    /// <summary>Optional per-row minimum heights in points. Null entries fall back to <see cref="MinRowHeight"/>.</summary>
    public System.Collections.Generic.List<double?>? RowMinHeights {
        get => _rowMinHeights;
        set {
            if (value == null) {
                _rowMinHeights = null;
                return;
            }

            var heights = new System.Collections.Generic.List<double?>(value.Count);
            foreach (double? height in value) {
                ValidateOptionalNonNegativeFiniteValue(height, nameof(RowMinHeights), "Table row minimum heights must be non-negative finite values.");
                heights.Add(height);
            }

            _rowMinHeights = heights;
        }
    }
    /// <summary>Optional per-row fixed heights in points. Null entries use content-driven row sizing and minimum row heights.</summary>
    public System.Collections.Generic.List<double?>? FixedRowHeights {
        get => _fixedRowHeights;
        set {
            if (value == null) {
                _fixedRowHeights = null;
                return;
            }

            var heights = new System.Collections.Generic.List<double?>(value.Count);
            foreach (double? height in value) {
                ValidateOptionalNonNegativeFiniteValue(height, nameof(FixedRowHeights), "Table fixed row heights must be non-negative finite values.");
                heights.Add(height);
            }

            _fixedRowHeights = heights;
        }
    }
    /// <summary>Vertical space before the table, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Table spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }
    /// <summary>Vertical space to reserve before table content when the same table continues on a new page.</summary>
    public double PageContinuationSpacingBefore {
        get => _pageContinuationSpacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(PageContinuationSpacingBefore), "Table page continuation spacing before must be a non-negative finite value.");
            _pageContinuationSpacingBefore = value;
        }
    }
    /// <summary>Optional text rendered above the table grid as part of the table flow.</summary>
    public string? Caption { get; set; }
    /// <summary>Optional alternate text attached to the tagged PDF table structure. This is not rendered visibly.</summary>
    public string? AlternativeText { get; set; }
    /// <summary>Caption alignment inside the rendered table width.</summary>
    public PdfAlign CaptionAlign {
        get => _captionAlign;
        set {
            Guard.LeftCenterRightAlign(value, nameof(CaptionAlign), "Table caption");
            _captionAlign = value;
        }
    }
    /// <summary>Caption text color. When null the writer's default text color is used.</summary>
    public PdfColor? CaptionColor { get; set; }
    /// <summary>Caption font size in points. When null the document default font size is used.</summary>
    public double? CaptionFontSize {
        get => _captionFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(CaptionFontSize), "Table caption font size must be a positive finite value.");
            _captionFontSize = value;
        }
    }
    /// <summary>Vertical space between the caption and table grid, in points.</summary>
    public double CaptionSpacingAfter {
        get => _captionSpacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(CaptionSpacingAfter), "Table caption spacing after must be a non-negative finite value.");
            _captionSpacingAfter = value;
        }
    }
    /// <summary>Vertical space after the table, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Table spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }
    /// <summary>Baseline adjustment in points to fine-tune text vertical placement.</summary>
    public double RowBaselineOffset {
        get => _rowBaselineOffset;
        set {
            ValidateFiniteValue(value, nameof(RowBaselineOffset), "Table row baseline offset must be a finite value.");
            _rowBaselineOffset = value;
        }
    }
    /// <summary>Optional maximum rendered table width, in points. When null the table uses the available frame width.</summary>
    public double? MaxWidth {
        get => _maxWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(MaxWidth), "Table max width must be a positive finite value.");
            _maxWidth = value;
        }
    }
    /// <summary>
    /// Optional preferred rendered table width, in points. Auto-fit tables may expand beyond this width
    /// to satisfy measured content, up to <see cref="MaxWidth"/> or the available frame width.
    /// </summary>
    public double? PreferredWidth {
        get => _preferredWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(PreferredWidth), "Table preferred width must be a positive finite value.");
            _preferredWidth = value;
        }
    }
    /// <summary>When true, the resolved table frame width is preserved even if measured columns would otherwise shrink to their content.</summary>
    public bool PreserveWidth { get; set; }
    /// <summary>When true, table rows can reduce their font size to keep cell text within the resolved cell width.</summary>
    public bool ShrinkTextToFit { get; set; }
    /// <summary>Smallest font size, in points, used by <see cref="ShrinkTextToFit"/>. When null the writer uses 6 points.</summary>
    public double? MinimumShrinkFontSize {
        get => _minimumShrinkFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(MinimumShrinkFontSize), "Table minimum shrink font size must be a positive finite value.");
            _minimumShrinkFontSize = value;
        }
    }
    /// <summary>Optional left indentation before table placement, in points.</summary>
    public double LeftIndent {
        get => _leftIndent;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(LeftIndent), "Table left indent must be a non-negative finite value.");
            _leftIndent = value;
        }
    }
    /// <summary>Optional per-column alignment; defaults to Left.</summary>
    public System.Collections.Generic.List<PdfColumnAlign>? Alignments {
        get => _alignments;
        set {
            if (value != null) {
                foreach (var alignment in value) {
                    Guard.TableColumnAlign(alignment, nameof(Alignments));
                }
            }

            _alignments = value == null ? null : new System.Collections.Generic.List<PdfColumnAlign>(value);
        }
    }
    /// <summary>Optional per-column vertical alignment; defaults to Top.</summary>
    public System.Collections.Generic.List<PdfCellVerticalAlign>? VerticalAlignments {
        get => _verticalAlignments;
        set {
            if (value != null) {
                foreach (var alignment in value) {
                    Guard.TableCellVerticalAlign(alignment, nameof(VerticalAlignments));
                }
            }

            _verticalAlignments = value == null ? null : new System.Collections.Generic.List<PdfCellVerticalAlign>(value);
        }
    }
    /// <summary>Optional fixed column widths in points. Null or missing entries are sized by relative weights.</summary>
    public System.Collections.Generic.List<double?>? ColumnWidthPoints {
        get => _columnWidthPoints;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnWidthPoints), "Table fixed column widths must be positive finite values.");
            _columnWidthPoints = value == null ? null : new System.Collections.Generic.List<double?>(value);
        }
    }
    /// <summary>Optional minimum column widths in points. Null or missing entries have no explicit minimum.</summary>
    public System.Collections.Generic.List<double?>? ColumnMinWidthPoints {
        get => _columnMinWidthPoints;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnMinWidthPoints), "Table minimum column widths must be positive finite values.");
            _columnMinWidthPoints = value == null ? null : new System.Collections.Generic.List<double?>(value);
        }
    }
    /// <summary>Optional maximum column widths in points. Null or missing entries have no explicit maximum.</summary>
    public System.Collections.Generic.List<double?>? ColumnMaxWidthPoints {
        get => _columnMaxWidthPoints;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnMaxWidthPoints), "Table maximum column widths must be positive finite values.");
            _columnMaxWidthPoints = value == null ? null : new System.Collections.Generic.List<double?>(value);
        }
    }
    /// <summary>Optional relative column width weights. Missing columns default to 1.0.</summary>
    public System.Collections.Generic.List<double>? ColumnWidthWeights {
        get => _columnWidthWeights;
        set {
            ValidatePositiveFiniteValues(value, nameof(ColumnWidthWeights), "Table column width weights must be positive finite values.");
            _columnWidthWeights = value == null ? null : new System.Collections.Generic.List<double>(value);
        }
    }
    /// <summary>When true, flexible column widths are weighted from measured table content using OfficeIMO.Drawing.</summary>
    public bool AutoFitColumns { get; set; }
    /// <summary>When true, cell values that look numeric are right-aligned automatically.</summary>
    public bool RightAlignNumeric { get; set; }
    /// <summary>When true, the table moves as a unit instead of splitting across pages when it fits in the page frame.</summary>
    public bool KeepTogether { get; set; }
    /// <summary>When true, the table moves with the first visible part of the following block when they fit together.</summary>
    public bool KeepWithNext { get; set; }
    /// <summary>When true, a single row that is taller than the page frame may split across pages by wrapped text line.</summary>
    public bool AllowRowBreakAcrossPages { get; set; } = true;
    /// <summary>Optional per-row row-break overrides. Null entries fall back to <see cref="AllowRowBreakAcrossPages"/>.</summary>
    public System.Collections.Generic.List<bool?>? RowAllowBreakAcrossPages {
        get => _rowAllowBreakAcrossPages;
        set => _rowAllowBreakAcrossPages = value == null ? null : new System.Collections.Generic.List<bool?>(value);
    }

    /// <summary>Creates a deep copy of this style.</summary>
    public PdfTableStyle Clone() {
        var clone = new PdfTableStyle {
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            HeaderFill = HeaderFill,
            FooterFill = FooterFill,
            RowStripeFill = RowStripeFill,
            RowSeparatorColor = RowSeparatorColor,
            RowSeparatorWidth = RowSeparatorWidth,
            HeaderSeparatorColor = HeaderSeparatorColor,
            HeaderSeparatorWidth = HeaderSeparatorWidth,
            FooterSeparatorColor = FooterSeparatorColor,
            FooterSeparatorWidth = FooterSeparatorWidth,
            TextColor = TextColor,
            FontSize = FontSize,
            LineHeight = LineHeight,
            HeaderTextColor = HeaderTextColor,
            HeaderFontSize = HeaderFontSize,
            HeaderBold = HeaderBold,
            HeaderRowCount = HeaderRowCount,
            RepeatHeaderRowCount = RepeatHeaderRowCount,
            FooterRowCount = FooterRowCount,
            MinimumBodyRowsOnLastPage = MinimumBodyRowsOnLastPage,
            FooterTextColor = FooterTextColor,
            FooterFontSize = FooterFontSize,
            FooterBold = FooterBold,
            CellPaddingX = CellPaddingX,
            CellPaddingY = CellPaddingY,
            CellPaddingLeft = CellPaddingLeft,
            CellPaddingRight = CellPaddingRight,
            CellPaddingTop = CellPaddingTop,
            CellPaddingBottom = CellPaddingBottom,
            CellSpacing = CellSpacing,
            MinRowHeight = MinRowHeight,
            RowMinHeights = RowMinHeights,
            FixedRowHeights = FixedRowHeights,
            SpacingBefore = SpacingBefore,
            PageContinuationSpacingBefore = PageContinuationSpacingBefore,
            Caption = Caption,
            AlternativeText = AlternativeText,
            CaptionAlign = CaptionAlign,
            CaptionColor = CaptionColor,
            CaptionFontSize = CaptionFontSize,
            CaptionSpacingAfter = CaptionSpacingAfter,
            SpacingAfter = SpacingAfter,
            RowBaselineOffset = RowBaselineOffset,
            PreferredWidth = PreferredWidth,
            MaxWidth = MaxWidth,
            PreserveWidth = PreserveWidth,
            ShrinkTextToFit = ShrinkTextToFit,
            MinimumShrinkFontSize = MinimumShrinkFontSize,
            LeftIndent = LeftIndent,
            AutoFitColumns = AutoFitColumns,
            RightAlignNumeric = RightAlignNumeric,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext,
            AllowRowBreakAcrossPages = AllowRowBreakAcrossPages,
            RowAllowBreakAcrossPages = RowAllowBreakAcrossPages
        };
        if (BodyColumnFills != null) clone.BodyColumnFills = new System.Collections.Generic.List<PdfColor?>(BodyColumnFills);
        if (CellFills != null) clone.CellFills = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor>(CellFills);
        if (CellDataBars != null) {
            clone.CellDataBars = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellDataBar>();
            foreach (var cellDataBar in CellDataBars) {
                clone.CellDataBars[cellDataBar.Key] = cellDataBar.Value.Clone();
            }
        }
        if (CellIcons != null) {
            clone.CellIcons = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellIcon>();
            foreach (var cellIcon in CellIcons) {
                clone.CellIcons[cellIcon.Key] = cellIcon.Value.Clone();
            }
        }
        if (CellBorders != null) {
            clone.CellBorders = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellBorder>();
            foreach (var cellBorder in CellBorders) {
                clone.CellBorders[cellBorder.Key] = cellBorder.Value.Clone();
            }
        }
        if (CellPaddings != null) {
            clone.CellPaddings = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>();
            foreach (var cellPadding in CellPaddings) {
                clone.CellPaddings[cellPadding.Key] = cellPadding.Value.Clone();
            }
        }
        if (CellAlignments != null) clone.CellAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>(CellAlignments);
        if (CellVerticalAlignments != null) clone.CellVerticalAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>(CellVerticalAlignments);
        if (Alignments != null) clone.Alignments = new System.Collections.Generic.List<PdfColumnAlign>(Alignments);
        if (VerticalAlignments != null) clone.VerticalAlignments = new System.Collections.Generic.List<PdfCellVerticalAlign>(VerticalAlignments);
        if (ColumnWidthPoints != null) clone.ColumnWidthPoints = new System.Collections.Generic.List<double?>(ColumnWidthPoints);
        if (ColumnMinWidthPoints != null) clone.ColumnMinWidthPoints = new System.Collections.Generic.List<double?>(ColumnMinWidthPoints);
        if (ColumnMaxWidthPoints != null) clone.ColumnMaxWidthPoints = new System.Collections.Generic.List<double?>(ColumnMaxWidthPoints);
        if (ColumnWidthWeights != null) clone.ColumnWidthWeights = new System.Collections.Generic.List<double>(ColumnWidthWeights);
        return clone;
    }

    private static void ValidateOptionalPositiveFiniteValues(System.Collections.Generic.IEnumerable<double?>? values, string paramName, string message) {
        if (values == null) {
            return;
        }

        foreach (var value in values) {
            if (value.HasValue && (value.Value <= 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
                throw new System.ArgumentException(message, paramName);
            }
        }
    }

    private static void ValidatePositiveFiniteValues(System.Collections.Generic.IEnumerable<double>? values, string paramName, string message) {
        if (values == null) {
            return;
        }

        foreach (double value in values) {
            if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException(message, paramName);
            }
        }
    }

    private static void ValidateNonNegativeValue(int value, string paramName, string message) {
        if (value < 0) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidatePositiveFiniteValue(double value, string paramName, string message) {
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateOptionalNonNegativeFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidateNonNegativeFiniteValue(value.Value, paramName, message);
        }
    }

    private static void ValidateOptionalPositiveFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidatePositiveFiniteValue(value.Value, paramName, message);
        }
    }

    private static void ValidateFiniteValue(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateCellBorder(PdfCellBorder? border, string paramName) {
        if (border == null || border.Width < 0 || double.IsNaN(border.Width) || double.IsInfinity(border.Width)) {
            throw new System.ArgumentException("Table cell border widths must be non-negative finite values.", paramName);
        }
    }

    private static void ValidateCellPadding(PdfCellPadding? padding, string paramName) {
        if (padding == null) {
            throw new System.ArgumentException("Table cell padding values cannot be null.", paramName);
        }
    }
}
