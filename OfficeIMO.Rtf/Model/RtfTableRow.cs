namespace OfficeIMO.Rtf;

/// <summary>
/// RTF table row.
/// </summary>
public sealed class RtfTableRow {
    private readonly List<RtfTableCell> _cells = new List<RtfTableCell>();

    /// <summary>Cells in the row.</summary>
    public IReadOnlyList<RtfTableCell> Cells => _cells.AsReadOnly();

    /// <summary>Whether this row repeats as a table header.</summary>
    public bool RepeatHeader { get; set; }

    /// <summary>Whether this row should be kept together on one page, represented by <c>\trkeep</c>.</summary>
    public bool KeepTogether { get; set; }

    /// <summary>Whether this row should be kept with the following row, represented by <c>\trkeepfollow</c>.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Optional AutoFit toggle for this row, represented by <c>\trautofit</c>.</summary>
    public bool? AutoFit { get; set; }

    /// <summary>Optional table row reading order, represented by <c>\ltrrow</c> or <c>\rtlrow</c>.</summary>
    public RtfTableRowDirection? Direction { get; set; }

    /// <summary>Preferred row height in twips.</summary>
    public int? HeightTwips { get; set; }

    /// <summary>Half of the inter-cell gap in twips represented by <c>\trgaph</c>.</summary>
    public int? CellGapTwips { get; set; }

    /// <summary>Left row indent in twips represented by <c>\trleft</c>.</summary>
    public int? LeftIndentTwips { get; set; }

    /// <summary>Optional row alignment represented by <c>\trql</c>, <c>\trqc</c>, or <c>\trqr</c>.</summary>
    public RtfTableAlignment? Alignment { get; set; }

    /// <summary>Preferred table width value carried by <c>\trwWidth</c> for this row definition.</summary>
    public int? PreferredWidth { get; set; }

    /// <summary>Preferred table width unit carried by <c>\trftsWidth</c> for this row definition.</summary>
    public RtfTableWidthUnit? PreferredWidthUnit { get; set; }

    /// <summary>One-based color table index used for table row background shading.</summary>
    public int? BackgroundColorIndex { get; set; }

    /// <summary>One-based color table index used for table row pattern foreground shading.</summary>
    public int? ShadingForegroundColorIndex { get; set; }

    /// <summary>Raw RTF <c>\trpat</c> table row shading pattern value.</summary>
    public int? ShadingPatternValue { get; set; }

    /// <summary>Raw RTF <c>\trshdng</c> value, where 10000 represents 100 percent.</summary>
    public int? ShadingPatternPercent { get; set; }

    /// <summary>Named RTF table row shading pattern.</summary>
    public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;

    /// <summary>Default top cell padding for this row in twips.</summary>
    public int? PaddingTopTwips { get; set; }

    /// <summary>Default left cell padding for this row in twips.</summary>
    public int? PaddingLeftTwips { get; set; }

    /// <summary>Default bottom cell padding for this row in twips.</summary>
    public int? PaddingBottomTwips { get; set; }

    /// <summary>Default right cell padding for this row in twips.</summary>
    public int? PaddingRightTwips { get; set; }

    /// <summary>Default top cell spacing for this row in twips.</summary>
    public int? SpacingTopTwips { get; set; }

    /// <summary>Default left cell spacing for this row in twips.</summary>
    public int? SpacingLeftTwips { get; set; }

    /// <summary>Default bottom cell spacing for this row in twips.</summary>
    public int? SpacingBottomTwips { get; set; }

    /// <summary>Default right cell spacing for this row in twips.</summary>
    public int? SpacingRightTwips { get; set; }

    /// <summary>Whether this positioned row should avoid overlapping similar floating tables, represented by <c>\tabsnoovrlp</c>.</summary>
    public bool NoOverlap { get; set; }

    /// <summary>Horizontal positioning reference frame for this row.</summary>
    public RtfTableHorizontalAnchor? HorizontalAnchor { get; set; }

    /// <summary>Vertical positioning reference frame for this row.</summary>
    public RtfTableVerticalAnchor? VerticalAnchor { get; set; }

    /// <summary>Horizontal table placement mode for this row.</summary>
    public RtfTableHorizontalPosition? HorizontalPosition { get; set; }

    /// <summary>Horizontal position value in twips for absolute placement modes.</summary>
    public int? HorizontalPositionTwips { get; set; }

    /// <summary>Vertical table placement mode for this row.</summary>
    public RtfTableVerticalPosition? VerticalPosition { get; set; }

    /// <summary>Vertical position value in twips for absolute placement modes.</summary>
    public int? VerticalPositionTwips { get; set; }

    /// <summary>Distance between surrounding text and the left edge of a positioned table.</summary>
    public int? TextWrapLeftTwips { get; set; }

    /// <summary>Distance between surrounding text and the right edge of a positioned table.</summary>
    public int? TextWrapRightTwips { get; set; }

    /// <summary>Distance between surrounding text and the top edge of a positioned table.</summary>
    public int? TextWrapTopTwips { get; set; }

    /// <summary>Distance between surrounding text and the bottom edge of a positioned table.</summary>
    public int? TextWrapBottomTwips { get; set; }

    /// <summary>Top row border.</summary>
    public RtfTableRowBorder TopBorder { get; } = new RtfTableRowBorder();

    /// <summary>Left row border.</summary>
    public RtfTableRowBorder LeftBorder { get; } = new RtfTableRowBorder();

    /// <summary>Bottom row border.</summary>
    public RtfTableRowBorder BottomBorder { get; } = new RtfTableRowBorder();

    /// <summary>Right row border.</summary>
    public RtfTableRowBorder RightBorder { get; } = new RtfTableRowBorder();

    /// <summary>Horizontal inside row border.</summary>
    public RtfTableRowBorder HorizontalBorder { get; } = new RtfTableRowBorder();

    /// <summary>Vertical inside row border.</summary>
    public RtfTableRowBorder VerticalBorder { get; } = new RtfTableRowBorder();

    /// <summary>Sets the row cell gap in twips.</summary>
    public RtfTableRow SetCellGap(int? cellGapTwips) {
        CellGapTwips = cellGapTwips;
        return this;
    }

    /// <summary>Sets the row left indent in twips.</summary>
    public RtfTableRow SetLeftIndent(int? leftIndentTwips) {
        LeftIndentTwips = leftIndentTwips;
        return this;
    }

    /// <summary>Sets the row alignment.</summary>
    public RtfTableRow SetAlignment(RtfTableAlignment? alignment) {
        Alignment = alignment;
        return this;
    }

    /// <summary>Sets the row AutoFit toggle.</summary>
    public RtfTableRow SetAutoFit(bool? autoFit) {
        AutoFit = autoFit;
        return this;
    }

    /// <summary>Sets the row reading order.</summary>
    public RtfTableRow SetDirection(RtfTableRowDirection? direction) {
        Direction = direction;
        return this;
    }

    /// <summary>Sets row background shading to a one-based color table index.</summary>
    public RtfTableRow SetBackgroundColor(int? colorIndex) {
        BackgroundColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets row shading color and pattern metadata.</summary>
    public RtfTableRow SetShading(
        int? backgroundColorIndex,
        int? foregroundColorIndex = null,
        int? patternValue = null,
        int? patternPercent = null,
        RtfShadingPattern pattern = RtfShadingPattern.None) {
        BackgroundColorIndex = backgroundColorIndex;
        ShadingForegroundColorIndex = foregroundColorIndex;
        ShadingPatternValue = patternValue;
        ShadingPatternPercent = patternPercent;
        ShadingPattern = pattern;
        return this;
    }

    /// <summary>Sets default row cell padding in twips.</summary>
    public RtfTableRow SetPadding(int? topTwips = null, int? leftTwips = null, int? bottomTwips = null, int? rightTwips = null) {
        PaddingTopTwips = topTwips;
        PaddingLeftTwips = leftTwips;
        PaddingBottomTwips = bottomTwips;
        PaddingRightTwips = rightTwips;
        return this;
    }

    /// <summary>Sets default row cell spacing in twips.</summary>
    public RtfTableRow SetSpacing(int? topTwips = null, int? leftTwips = null, int? bottomTwips = null, int? rightTwips = null) {
        SpacingTopTwips = topTwips;
        SpacingLeftTwips = leftTwips;
        SpacingBottomTwips = bottomTwips;
        SpacingRightTwips = rightTwips;
        return this;
    }

    /// <summary>Sets the positioned table row anchors.</summary>
    public RtfTableRow SetPositionAnchors(RtfTableHorizontalAnchor? horizontalAnchor, RtfTableVerticalAnchor? verticalAnchor) {
        HorizontalAnchor = horizontalAnchor;
        VerticalAnchor = verticalAnchor;
        return this;
    }

    /// <summary>Sets the positioned table row placement modes.</summary>
    public RtfTableRow SetPosition(
        RtfTableHorizontalPosition? horizontalPosition = null,
        int? horizontalTwips = null,
        RtfTableVerticalPosition? verticalPosition = null,
        int? verticalTwips = null) {
        HorizontalPosition = horizontalPosition;
        HorizontalPositionTwips = horizontalTwips;
        VerticalPosition = verticalPosition;
        VerticalPositionTwips = verticalTwips;
        return this;
    }

    /// <summary>Sets positioned table row text wrapping distances in twips.</summary>
    public RtfTableRow SetTextWrapDistances(int? leftTwips = null, int? rightTwips = null, int? topTwips = null, int? bottomTwips = null) {
        TextWrapLeftTwips = leftTwips;
        TextWrapRightTwips = rightTwips;
        TextWrapTopTwips = topTwips;
        TextWrapBottomTwips = bottomTwips;
        return this;
    }

    /// <summary>Adds a cell.</summary>
    public RtfTableCell AddCell(int? rightBoundaryTwips = null) {
        var cell = new RtfTableCell {
            RightBoundaryTwips = rightBoundaryTwips
        };
        _cells.Add(cell);
        return cell;
    }

    /// <summary>Gets one of the row borders.</summary>
    public RtfTableRowBorder GetBorder(RtfTableRowBorderSide side) {
        switch (side) {
            case RtfTableRowBorderSide.Top:
                return TopBorder;
            case RtfTableRowBorderSide.Left:
                return LeftBorder;
            case RtfTableRowBorderSide.Bottom:
                return BottomBorder;
            case RtfTableRowBorderSide.Right:
                return RightBorder;
            case RtfTableRowBorderSide.Horizontal:
                return HorizontalBorder;
            default:
                return VerticalBorder;
        }
    }
}
