namespace OfficeIMO.Rtf;

/// <summary>
/// RTF table cell.
/// </summary>
public sealed class RtfTableCell {
    private readonly List<RtfParagraph> _paragraphs = new List<RtfParagraph>();

    /// <summary>Paragraphs in the cell.</summary>
    public IReadOnlyList<RtfParagraph> Paragraphs => _paragraphs.AsReadOnly();

    /// <summary>Right cell boundary in twips.</summary>
    public int? RightBoundaryTwips { get; set; }

    /// <summary>Horizontal merge marker.</summary>
    public RtfTableCellMerge HorizontalMerge { get; set; }

    /// <summary>Vertical merge marker.</summary>
    public RtfTableCellMerge VerticalMerge { get; set; }

    /// <summary>One-based color table index used for cell background shading.</summary>
    public int? BackgroundColorIndex { get; set; }

    /// <summary>One-based color table index used for cell pattern foreground shading.</summary>
    public int? ShadingForegroundColorIndex { get; set; }

    /// <summary>Raw RTF <c>\clshdng</c> value, where 10000 represents 100 percent.</summary>
    public int? ShadingPatternPercent { get; set; }

    /// <summary>Named RTF cell shading pattern.</summary>
    public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;

    /// <summary>Vertical alignment for cell content.</summary>
    public RtfTableCellVerticalAlignment? VerticalAlignment { get; set; }

    /// <summary>Text flow for cell content.</summary>
    public RtfTableCellTextFlow? TextFlow { get; set; }

    /// <summary>Preferred cell width value carried by <c>\clwWidth</c>.</summary>
    public int? PreferredWidth { get; set; }

    /// <summary>Preferred cell width unit carried by <c>\clftsWidth</c>.</summary>
    public RtfTableWidthUnit? PreferredWidthUnit { get; set; }

    /// <summary>Whether the end-of-cell mark is hidden, represented by <c>\clhidemark</c>.</summary>
    public bool HideCellMark { get; set; }

    /// <summary>Whether text should not wrap inside the cell, represented by <c>\clNoWrap</c>.</summary>
    public bool NoWrap { get; set; }

    /// <summary>Whether text is compressed to fit within the cell width, represented by <c>\clFitText</c>.</summary>
    public bool FitText { get; set; }

    /// <summary>Top cell padding in twips.</summary>
    public int? PaddingTopTwips { get; set; }

    /// <summary>Left cell padding in twips.</summary>
    public int? PaddingLeftTwips { get; set; }

    /// <summary>Bottom cell padding in twips.</summary>
    public int? PaddingBottomTwips { get; set; }

    /// <summary>Right cell padding in twips.</summary>
    public int? PaddingRightTwips { get; set; }

    /// <summary>Top cell border.</summary>
    public RtfTableCellBorder TopBorder { get; } = new RtfTableCellBorder();

    /// <summary>Left cell border.</summary>
    public RtfTableCellBorder LeftBorder { get; } = new RtfTableCellBorder();

    /// <summary>Bottom cell border.</summary>
    public RtfTableCellBorder BottomBorder { get; } = new RtfTableCellBorder();

    /// <summary>Right cell border.</summary>
    public RtfTableCellBorder RightBorder { get; } = new RtfTableCellBorder();

    /// <summary>Diagonal cell border from top-left to bottom-right.</summary>
    public RtfTableCellBorder TopLeftToBottomRightBorder { get; } = new RtfTableCellBorder();

    /// <summary>Diagonal cell border from top-right to bottom-left.</summary>
    public RtfTableCellBorder TopRightToBottomLeftBorder { get; } = new RtfTableCellBorder();

    /// <summary>Sets cell padding in twips.</summary>
    public RtfTableCell SetPadding(int? topTwips = null, int? leftTwips = null, int? bottomTwips = null, int? rightTwips = null) {
        PaddingTopTwips = topTwips;
        PaddingLeftTwips = leftTwips;
        PaddingBottomTwips = bottomTwips;
        PaddingRightTwips = rightTwips;
        return this;
    }

    /// <summary>Sets cell background shading to a one-based color table index.</summary>
    public RtfTableCell SetBackgroundColor(int? colorIndex) {
        BackgroundColorIndex = colorIndex;
        return this;
    }

    /// <summary>Sets cell shading color and pattern metadata.</summary>
    public RtfTableCell SetShading(int? backgroundColorIndex, int? foregroundColorIndex = null, int? patternPercent = null, RtfShadingPattern pattern = RtfShadingPattern.None) {
        BackgroundColorIndex = backgroundColorIndex;
        ShadingForegroundColorIndex = foregroundColorIndex;
        ShadingPatternPercent = patternPercent;
        ShadingPattern = pattern;
        return this;
    }

    /// <summary>Sets the cell text flow.</summary>
    public RtfTableCell SetTextFlow(RtfTableCellTextFlow? textFlow) {
        TextFlow = textFlow;
        return this;
    }

    /// <summary>Sets the preferred cell width.</summary>
    public RtfTableCell SetPreferredWidth(int? width, RtfTableWidthUnit? unit = null) {
        PreferredWidth = width;
        PreferredWidthUnit = unit;
        return this;
    }

    /// <summary>Sets whether the end-of-cell mark is hidden.</summary>
    public RtfTableCell SetHideCellMark(bool hidden = true) {
        HideCellMark = hidden;
        return this;
    }

    /// <summary>Sets whether text should not wrap inside the cell.</summary>
    public RtfTableCell SetNoWrap(bool noWrap = true) {
        NoWrap = noWrap;
        return this;
    }

    /// <summary>Sets whether text is compressed to fit within the cell width.</summary>
    public RtfTableCell SetFitText(bool fitText = true) {
        FitText = fitText;
        return this;
    }

    /// <summary>Adds a paragraph to the cell.</summary>
    public RtfParagraph AddParagraph(string? text = null) {
        var paragraph = new RtfParagraph();
        if (!string.IsNullOrEmpty(text)) {
            paragraph.AddText(text!);
        }

        _paragraphs.Add(paragraph);
        return paragraph;
    }

    internal void AddParsedParagraph(RtfParagraph paragraph) {
        _paragraphs.Add(paragraph ?? throw new ArgumentNullException(nameof(paragraph)));
    }
}
