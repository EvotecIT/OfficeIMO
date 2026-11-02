namespace OfficeIMO.Pdf;

/// <summary>
/// Describes visual and layout options for table rendering.
/// Attach an instance to a table block or use the presets in <see cref="PdfTableStyles"/>.
/// </summary>
public partial class PdfTableStyle {
    private PdfAlign _captionAlign = PdfAlign.Left;
    private System.Collections.Generic.List<PdfColumnAlign>? _alignments;
    private System.Collections.Generic.List<PdfCellVerticalAlign>? _verticalAlignments;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? _columnWidths;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? _columnMinWidths;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? _columnMaxWidths;
    private System.Collections.Generic.List<double>? _columnWidthWeights;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeColor?>? _bodyColumnFills;
    private System.Collections.Generic.Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>? _cellFills;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellDataBar>? _cellDataBars;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellIcon>? _cellIcons;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellBorder>? _cellBorders;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>? _cellPaddings;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>? _cellAlignments;
    private System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>? _cellVerticalAlignments;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? _rowMinHeights;
    private System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? _fixedRowHeights;
    private System.Collections.Generic.List<bool?>? _rowAllowBreakAcrossPages;
    private global::OfficeIMO.OfficeLength _borderWidth = global::OfficeIMO.OfficeLength.Points(0.5);
    private global::OfficeIMO.OfficeLength _cornerRadius;
    private global::OfficeIMO.OfficeLength _rowSeparatorWidth;
    private global::OfficeIMO.OfficeLength _headerSeparatorWidth;
    private global::OfficeIMO.OfficeLength _footerSeparatorWidth;
    private int _headerRowCount = 1;
    private int _footerRowCount;
    private int? _repeatHeaderRowCount;
    private int _minimumBodyRowsOnFirstPage = 2;
    private int _minimumBodyRowsOnLastPage = 2;
    private global::OfficeIMO.OfficeLength _cellPaddingX = global::OfficeIMO.OfficeLength.Points(4);
    private global::OfficeIMO.OfficeLength _cellPaddingY = global::OfficeIMO.OfficeLength.Points(2);
    private global::OfficeIMO.OfficeLength? _cellPaddingLeft;
    private global::OfficeIMO.OfficeLength? _cellPaddingRight;
    private global::OfficeIMO.OfficeLength? _cellPaddingTop;
    private global::OfficeIMO.OfficeLength? _cellPaddingBottom;
    private global::OfficeIMO.OfficeLength _cellSpacing;
    private global::OfficeIMO.OfficeLength _minRowHeight;
    private global::OfficeIMO.OfficeLength _spacingBefore;
    private global::OfficeIMO.OfficeLength _pageContinuationSpacingBefore;
    private global::OfficeIMO.OfficeLength? _captionFontSize;
    private global::OfficeIMO.OfficeLength _captionSpacingAfter = global::OfficeIMO.OfficeLength.Points(4);
    private global::OfficeIMO.OfficeLength _spacingAfter;
    private global::OfficeIMO.OfficeLength _rowBaselineOffset;
    private global::OfficeIMO.OfficeLength? _preferredWidth;
    private global::OfficeIMO.OfficeLength? _maxWidth;
    private global::OfficeIMO.OfficeLength? _fontSize;
    private double? _lineHeight;
    private global::OfficeIMO.OfficeLength? _headerFontSize;
    private global::OfficeIMO.OfficeLength? _footerFontSize;
    private global::OfficeIMO.OfficeLength? _minimumShrinkFontSize;

    /// <summary>Color of the table borders and cell grid lines. Set to null to hide borders.</summary>
    public global::OfficeIMO.OfficeColor? BorderColor { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.8, 0.8, 0.8);
    /// <summary>
    /// Whether the table advances the vertical flow cursor. Set to false for a positioned table
    /// that overlays the following flow; the table is still drawn at the current cursor position.
    /// </summary>
    public bool ConsumesVerticalFlow { get; set; } = true;
    /// <summary>Document adapters can retain glyph positions while containing their painting in the cell.</summary>
    internal bool ClipTextToCellBounds { get; set; }
    /// <summary>Drops the horizontal padding of a column narrower than it, as Word does for thin spacer columns, instead of rejecting the table.</summary>
    internal bool DropPaddingInNarrowColumns { get; set; }
    internal bool PreservePartialCellLines { get; set; }
    /// <summary>
    /// Document adapters can measure cell margins from the inner edge of the border
    /// paint and retain clearance from neighboring cell borders.
    /// </summary>
    internal bool CellPaddingFromBorderInterior { get; set; }
    /// <summary>Resolves source cell rules against inherited rules at physical row boundaries.</summary>
    internal bool CollapseHorizontalCellBorders { get; set; }
    /// <summary>Prepared inside-frame cells retain the vertical margin clip after their border insets are resolved.</summary>
    internal bool CellTextClipIncludesVerticalPadding { get; set; }
    /// <summary>Imported fixed grids compress the width above each physical cell minimum.</summary>
    internal bool CompressFixedColumnsAboveMinimums { get; set; }

    /// <summary>Document adapters can retain a perimeter independent of the cell borders.</summary>
    internal PdfTableBorderFrame? BorderFrame { get; set; }

    /// <summary>
    /// Expands an authored automatic table width only for content that cannot
    /// wrap within the preferred grid, rather than for the unwrapped line width.
    /// </summary>
    internal bool AutoFitWidthUsesContentMinimum { get; set; }
    /// <summary>
    /// Captured absolute source preferences remain the compression basis when
    /// their complete grid exceeds an explicit width limit. Normal AutoFit
    /// content allocation remains active while that source budget fits.
    /// </summary>
    internal bool PreserveAuthoredPreferencesAtWidthLimit { get; set; }
    /// <summary>Document adapters can size an explicit automatic grid that has no positive authored width from its content minimum.</summary>
    internal bool AutoFitUnspecifiedWidthToContent { get; set; }

    // Source-relative constraints are resolved only after the containing cell's
    // actual inner frame has been selected by the shared column owner.
    internal double? PreferredWidthFraction { get; set; }
    internal double? MaxWidthFraction { get; set; }
    internal double FrameWidthAdjustment { get; set; }
    /// <summary>
    /// Converts a resolved source-relative frame width into grid width after
    /// percentage scaling, for example when the source width includes its own spacing.
    /// </summary>
    internal double ResolvedWidthAdjustment { get; set; }
    /// <summary>Authored automatic-column preferences remain distinct from their measured content minimums.</summary>
    internal System.Collections.Generic.List<double?>? AutoFitColumnPreferredWidthPoints { get; set; }
    /// <summary>Source cells whose own absolute preference permits wrapped content instead of an unwrapped span preference.</summary>
    internal System.Collections.Generic.Dictionary<(int Row, int Column), double>? AutoFitCellPreferredWidthPoints { get; set; }
    /// <summary>Original imported percentage preferences; null columns share the remaining width.</summary>
    internal System.Collections.Generic.List<double?>? ImportedColumnWidthFractions { get; set; }
    /// <summary>Optional floating placement. A positioned table reserves space for surrounding text without advancing the flow cursor.</summary>
    public PdfTablePosition? Position { get; set; }
    /// <summary>Stroke width, as a physical length, for table borders and cell grid lines.</summary>
    public global::OfficeIMO.OfficeLength BorderWidth {
        get => _borderWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(BorderWidth), "Table border width must be a non-negative finite value.");
            _borderWidth = value;
        }
    }
    /// <summary>
    /// Corner radius, as a physical length, for the table's outer box. Zero (the default) keeps square corners.
    /// The rounding is applied to the outer perimeter only: the top corners of the first row and the
    /// bottom corners of the last row are rounded, while interior grid lines stay straight. Cell fills
    /// and per-cell border overrides (for example an accent stripe or a top-accent bar) are clipped to
    /// the rounded outer box so they follow the curve instead of overrunning it. Most useful for
    /// single-row / single-cell tables such as callouts and stat cards.
    /// </summary>
    public global::OfficeIMO.OfficeLength CornerRadius {
        get => _cornerRadius;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CornerRadius), "Table corner radius must be a non-negative finite value.");
            _cornerRadius = value;
        }
    }
    /// <summary>Background fill color for the header row. Set to null for no fill.</summary>
    public global::OfficeIMO.OfficeColor? HeaderFill { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.95, 0.95, 0.95);
    /// <summary>Background fill color for footer rows. Set to null for no fill.</summary>
    public global::OfficeIMO.OfficeColor? FooterFill { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.95, 0.95, 0.95);
    /// <summary>Optional alternating row fill color (applied to every other body row). Set to null to disable.</summary>
    public global::OfficeIMO.OfficeColor? RowStripeFill { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.98, 0.98, 0.98);
    /// <summary>Optional horizontal separator color drawn at the bottom of each table row. Set to null or zero width to disable.</summary>
    public global::OfficeIMO.OfficeColor? RowSeparatorColor { get; set; }
    /// <summary>Stroke width, as a physical length, for row separators.</summary>
    public global::OfficeIMO.OfficeLength RowSeparatorWidth {
        get => _rowSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(RowSeparatorWidth), "Table row separator width must be a non-negative finite value.");
            _rowSeparatorWidth = value;
        }
    }
    /// <summary>Optional horizontal separator color used below header rows. Falls back to <see cref="RowSeparatorColor"/> when null.</summary>
    public global::OfficeIMO.OfficeColor? HeaderSeparatorColor { get; set; }
    /// <summary>Optional stroke width, as a physical length, used below header rows. Falls back to <see cref="RowSeparatorWidth"/> when zero.</summary>
    public global::OfficeIMO.OfficeLength HeaderSeparatorWidth {
        get => _headerSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(HeaderSeparatorWidth), "Table header separator width must be a non-negative finite value.");
            _headerSeparatorWidth = value;
        }
    }
    /// <summary>Optional horizontal separator color used above footer rows. Falls back to <see cref="RowSeparatorColor"/> when null.</summary>
    public global::OfficeIMO.OfficeColor? FooterSeparatorColor { get; set; }
    /// <summary>Optional stroke width, as a physical length, used above footer rows. Falls back to <see cref="RowSeparatorWidth"/> when zero.</summary>
    public global::OfficeIMO.OfficeLength FooterSeparatorWidth {
        get => _footerSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(FooterSeparatorWidth), "Table footer separator width must be a non-negative finite value.");
            _footerSeparatorWidth = value;
        }
    }
    /// <summary>Optional per-column background fills for body rows. Null or missing entries use the row fill.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeColor?>? BodyColumnFills {
        get => _bodyColumnFills;
        set => _bodyColumnFills = value == null ? null : new System.Collections.Generic.List<global::OfficeIMO.OfficeColor?>(value);
    }
    /// <summary>Optional absolute per-cell background fills keyed by zero-based row and column.</summary>
    public System.Collections.Generic.Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>? CellFills {
        get => _cellFills;
        set {
            if (value != null) {
                foreach (var cellFill in value) {
                    if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                        throw new System.ArgumentException("Table cell fill coordinates cannot be negative.", nameof(CellFills));
                    }
                }
            }

            _cellFills = value == null ? null : new System.Collections.Generic.Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>(value);
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
    public global::OfficeIMO.OfficeColor? TextColor { get; set; }
    /// <summary>Font size for body cells, as a physical length. When null the document default font size is used.</summary>
    public global::OfficeIMO.OfficeLength? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(FontSize), "Table body font size must be a positive finite value.");
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
    public global::OfficeIMO.OfficeColor? HeaderTextColor { get; set; }
    /// <summary>Font size for header cells, as a physical length. When null the body table font size is used.</summary>
    public global::OfficeIMO.OfficeLength? HeaderFontSize {
        get => _headerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(HeaderFontSize), "Table header font size must be a positive finite value.");
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
    /// <summary>
    /// Minimum number of body rows kept with leading header rows when a table starts on a partially occupied page and the group fits on one page.
    /// Set to zero to allow a header-only first segment. Defaults to two.
    /// </summary>
    public int MinimumBodyRowsOnFirstPage {
        get => _minimumBodyRowsOnFirstPage;
        set {
            ValidateNonNegativeValue(value, nameof(MinimumBodyRowsOnFirstPage), "Table minimum first-page body row count cannot be negative.");
            _minimumBodyRowsOnFirstPage = value;
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
    public global::OfficeIMO.OfficeColor? FooterTextColor { get; set; }
    /// <summary>Font size for footer cells, as a physical length. When null the body table font size is used.</summary>
    public global::OfficeIMO.OfficeLength? FooterFontSize {
        get => _footerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(FooterFontSize), "Table footer font size must be a positive finite value.");
            _footerFontSize = value;
        }
    }
    /// <summary>When true, footer cells use the bold variant of the document font family.</summary>
    public bool FooterBold { get; set; } = true;
    /// <summary>Horizontal padding inside each cell, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength CellPaddingX {
        get => _cellPaddingX;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CellPaddingX), "Table horizontal cell padding must be a non-negative finite value.");
            _cellPaddingX = value;
        }
    }
    /// <summary>Vertical padding inside each cell, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength CellPaddingY {
        get => _cellPaddingY;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CellPaddingY), "Table vertical cell padding must be a non-negative finite value.");
            _cellPaddingY = value;
        }
    }
    /// <summary>Optional left padding override inside each cell, as a physical length. When null <see cref="CellPaddingX"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? CellPaddingLeft {
        get => _cellPaddingLeft;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(CellPaddingLeft), "Table left cell padding must be a non-negative finite value.");
            _cellPaddingLeft = value;
        }
    }
    /// <summary>Optional right padding override inside each cell, as a physical length. When null <see cref="CellPaddingX"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? CellPaddingRight {
        get => _cellPaddingRight;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(CellPaddingRight), "Table right cell padding must be a non-negative finite value.");
            _cellPaddingRight = value;
        }
    }
    /// <summary>Optional top padding override inside each cell, as a physical length. When null <see cref="CellPaddingY"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? CellPaddingTop {
        get => _cellPaddingTop;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(CellPaddingTop), "Table top cell padding must be a non-negative finite value.");
            _cellPaddingTop = value;
        }
    }
    /// <summary>Optional bottom padding override inside each cell, as a physical length. When null <see cref="CellPaddingY"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? CellPaddingBottom {
        get => _cellPaddingBottom;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(CellPaddingBottom), "Table bottom cell padding must be a non-negative finite value.");
            _cellPaddingBottom = value;
        }
    }
    /// <summary>Optional spacing between adjacent table cells, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength CellSpacing {
        get => _cellSpacing;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CellSpacing), "Table cell spacing must be a non-negative finite value.");
            _cellSpacing = value;
        }
    }
    /// <summary>Optional minimum row height as a physical length. Set to 0 to size rows from wrapped content.</summary>
    public global::OfficeIMO.OfficeLength MinRowHeight {
        get => _minRowHeight;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(MinRowHeight), "Table minimum row height must be a non-negative finite value.");
            _minRowHeight = value;
        }
    }
    /// <summary>Optional per-row minimum heights as a physical length. Null entries fall back to <see cref="MinRowHeight"/>.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? RowMinHeights {
        get => _rowMinHeights;
        set {
            if (value == null) {
                _rowMinHeights = null;
                return;
            }

            var heights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(value.Count);
            foreach (global::OfficeIMO.OfficeLength? height in value) {
                ValidateOptionalNonNegativeFiniteValue(height?.ToPoints(), nameof(RowMinHeights), "Table row minimum heights must be non-negative finite values.");
                heights.Add(height);
            }

            _rowMinHeights = heights;
        }
    }
    /// <summary>Optional per-row fixed heights as a physical length. Null entries use content-driven row sizing and minimum row heights.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? FixedRowHeights {
        get => _fixedRowHeights;
        set {
            if (value == null) {
                _fixedRowHeights = null;
                return;
            }

            var heights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(value.Count);
            foreach (global::OfficeIMO.OfficeLength? height in value) {
                ValidateOptionalNonNegativeFiniteValue(height?.ToPoints(), nameof(FixedRowHeights), "Table fixed row heights must be non-negative finite values.");
                heights.Add(height);
            }

            _fixedRowHeights = heights;
        }
    }
    /// <summary>Vertical space before the table, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(SpacingBefore), "Table spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }
    /// <summary>Vertical space to reserve before table content when the same table continues on a new page.</summary>
    public global::OfficeIMO.OfficeLength PageContinuationSpacingBefore {
        get => _pageContinuationSpacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(PageContinuationSpacingBefore), "Table page continuation spacing before must be a non-negative finite value.");
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
    public global::OfficeIMO.OfficeColor? CaptionColor { get; set; }
    /// <summary>Caption font size as a physical length. When null the document default font size is used.</summary>
    public global::OfficeIMO.OfficeLength? CaptionFontSize {
        get => _captionFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(CaptionFontSize), "Table caption font size must be a positive finite value.");
            _captionFontSize = value;
        }
    }
    /// <summary>Vertical space between the caption and table grid, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength CaptionSpacingAfter {
        get => _captionSpacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CaptionSpacingAfter), "Table caption spacing after must be a non-negative finite value.");
            _captionSpacingAfter = value;
        }
    }
    /// <summary>Vertical space after the table, as a physical length.</summary>
    public global::OfficeIMO.OfficeLength SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(SpacingAfter), "Table spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }
    /// <summary>Baseline adjustment as a physical length to fine-tune text vertical placement.</summary>
    public global::OfficeIMO.OfficeLength RowBaselineOffset {
        get => _rowBaselineOffset;
        set {
            ValidateFiniteValue(value.ToPoints(), nameof(RowBaselineOffset), "Table row baseline offset must be a finite value.");
            _rowBaselineOffset = value;
        }
    }
    /// <summary>Optional maximum rendered table width, as a physical length. When null the table uses the available frame width.</summary>
    public global::OfficeIMO.OfficeLength? MaxWidth {
        get => _maxWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(MaxWidth), "Table max width must be a positive finite value.");
            _maxWidth = value;
        }
    }
    /// <summary>
    /// Optional preferred rendered table width, as a physical length. Auto-fit tables may expand beyond this width
    /// to satisfy measured content, up to <see cref="MaxWidth"/> or the available frame width.
    /// </summary>
    public global::OfficeIMO.OfficeLength? PreferredWidth {
        get => _preferredWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(PreferredWidth), "Table preferred width must be a positive finite value.");
            _preferredWidth = value;
        }
    }
    /// <summary>When true, the resolved table frame width is preserved even if measured columns would otherwise shrink to their content.</summary>
    public bool PreserveWidth { get; set; }
    /// <summary>When true, table rows can reduce their font size to keep cell text within the resolved cell width.</summary>
    public bool ShrinkTextToFit { get; set; }
    /// <summary>Smallest font size, as a physical length, used by row or cell text shrinking. When null the writer uses 6 points.</summary>
    public global::OfficeIMO.OfficeLength? MinimumShrinkFontSize {
        get => _minimumShrinkFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(MinimumShrinkFontSize), "Table minimum shrink font size must be a positive finite value.");
            _minimumShrinkFontSize = value;
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
    /// <summary>Optional fixed column widths as a physical length. Null or missing entries are sized by relative weights.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? ColumnWidths {
        get => _columnWidths;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnWidths), "Table fixed column widths must be positive finite values.");
            _columnWidths = value == null ? null : new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(value);
        }
    }
    /// <summary>Optional minimum column widths as a physical length. Null or missing entries have no explicit minimum.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? ColumnMinWidths {
        get => _columnMinWidths;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnMinWidths), "Table minimum column widths must be positive finite values.");
            _columnMinWidths = value == null ? null : new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(value);
        }
    }
    /// <summary>Optional maximum column widths as a physical length. Null or missing entries have no explicit maximum.</summary>
    public System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>? ColumnMaxWidths {
        get => _columnMaxWidths;
        set {
            ValidateOptionalPositiveFiniteValues(value, nameof(ColumnMaxWidths), "Table maximum column widths must be positive finite values.");
            _columnMaxWidths = value == null ? null : new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(value);
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
            ConsumesVerticalFlow = ConsumesVerticalFlow,
            ClipTextToCellBounds = ClipTextToCellBounds,
            PreservePartialCellLines = PreservePartialCellLines,
            DropPaddingInNarrowColumns = DropPaddingInNarrowColumns,
            Position = Position,
            BorderWidth = BorderWidth,
            CornerRadius = CornerRadius,
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
            MinimumBodyRowsOnFirstPage = MinimumBodyRowsOnFirstPage,
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
            CellPaddingFromBorderInterior = CellPaddingFromBorderInterior,
            CollapseHorizontalCellBorders = CollapseHorizontalCellBorders,
            CellTextClipIncludesVerticalPadding = CellTextClipIncludesVerticalPadding,
            CompressFixedColumnsAboveMinimums = CompressFixedColumnsAboveMinimums,
            BorderFrame = BorderFrame?.Clone(),
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
            PreferredWidthFraction = PreferredWidthFraction,
            MaxWidthFraction = MaxWidthFraction,
            FrameWidthAdjustment = FrameWidthAdjustment,
            ResolvedWidthAdjustment = ResolvedWidthAdjustment,
            PreserveWidth = PreserveWidth,
            AllowWidthBeyondFrame = AllowWidthBeyondFrame,
            ShrinkTextToFit = ShrinkTextToFit,
            MinimumShrinkFontSize = MinimumShrinkFontSize,
            LeftIndent = LeftIndent,
            HorizontalOffset = HorizontalOffset,
            AutoFitColumns = AutoFitColumns,
            AutoFitWidthUsesContentMinimum = AutoFitWidthUsesContentMinimum,
            PreserveAuthoredPreferencesAtWidthLimit = PreserveAuthoredPreferencesAtWidthLimit,
            AutoFitUnspecifiedWidthToContent = AutoFitUnspecifiedWidthToContent,
            RightAlignNumeric = RightAlignNumeric,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext,
            AllowRowBreakAcrossPages = AllowRowBreakAcrossPages,
            RowAllowBreakAcrossPages = RowAllowBreakAcrossPages
        };
        if (BodyColumnFills != null) clone.BodyColumnFills = new System.Collections.Generic.List<global::OfficeIMO.OfficeColor?>(BodyColumnFills);
        if (AutoFitCellPreferredWidthPoints != null)
            clone.AutoFitCellPreferredWidthPoints = new System.Collections.Generic.Dictionary<(int Row, int Column), double>(AutoFitCellPreferredWidthPoints);
        if (CellFills != null) clone.CellFills = new System.Collections.Generic.Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>(CellFills);
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
        clone.CellHeaderScopes = CellHeaderScopes;
        if (CellAlignments != null) clone.CellAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign>(CellAlignments);
        if (CellVerticalAlignments != null) clone.CellVerticalAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign>(CellVerticalAlignments);
        if (Alignments != null) clone.Alignments = new System.Collections.Generic.List<PdfColumnAlign>(Alignments);
        if (VerticalAlignments != null) clone.VerticalAlignments = new System.Collections.Generic.List<PdfCellVerticalAlign>(VerticalAlignments);
        if (ColumnWidths != null) clone.ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(ColumnWidths);
        if (ColumnMinWidths != null) clone.ColumnMinWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(ColumnMinWidths);
        if (AutoFitColumnPreferredWidthPoints != null) clone.AutoFitColumnPreferredWidthPoints = new System.Collections.Generic.List<double?>(AutoFitColumnPreferredWidthPoints);
        if (ImportedColumnWidthFractions != null) clone.ImportedColumnWidthFractions = new System.Collections.Generic.List<double?>(ImportedColumnWidthFractions);
        if (ColumnMaxWidths != null) clone.ColumnMaxWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?>(ColumnMaxWidths);
        if (ColumnWidthWeights != null) clone.ColumnWidthWeights = new System.Collections.Generic.List<double>(ColumnWidthWeights);
        return clone;
    }

    private static void ValidateOptionalPositiveFiniteValues(System.Collections.Generic.IEnumerable<global::OfficeIMO.OfficeLength?>? values, string paramName, string message) {
        if (values == null) {
            return;
        }

        foreach (var value in values) {
            if (value.HasValue && (value.Value.ToPoints() <= 0 || double.IsNaN(value.Value.ToPoints()) || double.IsInfinity(value.Value.ToPoints()))) {
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
        if (border == null || border.Width.ToPoints() < 0 || double.IsNaN(border.Width.ToPoints()) || double.IsInfinity(border.Width.ToPoints())) {
            throw new System.ArgumentException("Table cell border widths must be non-negative finite values.", paramName);
        }
    }

    private static void ValidateCellPadding(PdfCellPadding? padding, string paramName) {
        if (padding == null) {
            throw new System.ArgumentException("Table cell padding values cannot be null.", paramName);
        }
    }
}
