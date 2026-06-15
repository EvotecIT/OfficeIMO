namespace OfficeIMO.Rtf;

/// <summary>
/// Page size and margin settings for an RTF document.
/// </summary>
public sealed class RtfPageSetup {
    /// <summary>Paper width in twips.</summary>
    public int? PaperWidthTwips { get; set; }

    /// <summary>Paper height in twips.</summary>
    public int? PaperHeightTwips { get; set; }

    /// <summary>Printer paper size code represented by <c>\psz</c>.</summary>
    public int? PrinterPaperSize { get; set; }

    /// <summary>Printer bin source for the first page represented by <c>\binfsxn</c>.</summary>
    public int? FirstPagePaperSource { get; set; }

    /// <summary>Printer bin source for subsequent pages represented by <c>\binsxn</c>.</summary>
    public int? OtherPagesPaperSource { get; set; }

    /// <summary>Left margin in twips.</summary>
    public int? MarginLeftTwips { get; set; }

    /// <summary>Right margin in twips.</summary>
    public int? MarginRightTwips { get; set; }

    /// <summary>Top margin in twips.</summary>
    public int? MarginTopTwips { get; set; }

    /// <summary>Bottom margin in twips.</summary>
    public int? MarginBottomTwips { get; set; }

    /// <summary>Gutter width in twips.</summary>
    public int? GutterWidthTwips { get; set; }

    /// <summary>Header distance from the top edge of the page in twips.</summary>
    public int? HeaderDistanceTwips { get; set; }

    /// <summary>Footer distance from the bottom edge of the page in twips.</summary>
    public int? FooterDistanceTwips { get; set; }

    /// <summary>Beginning page number.</summary>
    public int? PageNumberStart { get; set; }

    /// <summary>Whether page numbering restarts in this page setup. <c>false</c> represents continuous numbering.</summary>
    public bool? PageNumberRestart { get; set; }

    /// <summary>Horizontal page-number position in twips from the right margin.</summary>
    public int? PageNumberPositionXTwips { get; set; }

    /// <summary>Vertical page-number position in twips from the top margin.</summary>
    public int? PageNumberPositionYTwips { get; set; }

    /// <summary>Page-number display format.</summary>
    public RtfPageNumberFormat? PageNumberFormat { get; set; }

    /// <summary>Page border settings.</summary>
    public RtfPageBorders PageBorders { get; } = new RtfPageBorders();

    /// <summary>Whether the document declares landscape orientation.</summary>
    public bool Landscape { get; set; }

    /// <summary>Whether first-page header/footer variants are enabled, represented by <c>\titlepg</c>.</summary>
    public bool DifferentFirstPageHeaderFooter { get; set; }

    /// <summary>Whether the gutter should be positioned on the right side.</summary>
    public bool RtlGutter { get; set; }

    /// <summary>Sets paper size in twips.</summary>
    public RtfPageSetup SetPaperSize(int widthTwips, int heightTwips) {
        if (widthTwips <= 0) throw new ArgumentOutOfRangeException(nameof(widthTwips), "Paper width must be greater than zero.");
        if (heightTwips <= 0) throw new ArgumentOutOfRangeException(nameof(heightTwips), "Paper height must be greater than zero.");
        PaperWidthTwips = widthTwips;
        PaperHeightTwips = heightTwips;
        return this;
    }

    /// <summary>Sets printer-specific paper metadata.</summary>
    public RtfPageSetup SetPrinterPaper(int? paperSize = null, int? firstPageSource = null, int? otherPagesSource = null) {
        ValidateNonNegative(paperSize, nameof(paperSize));
        ValidateNonNegative(firstPageSource, nameof(firstPageSource));
        ValidateNonNegative(otherPagesSource, nameof(otherPagesSource));
        PrinterPaperSize = paperSize;
        FirstPagePaperSource = firstPageSource;
        OtherPagesPaperSource = otherPagesSource;
        return this;
    }

    /// <summary>Sets gutter width in twips.</summary>
    public RtfPageSetup SetGutter(int? gutterWidthTwips = null, bool rtlGutter = false) {
        ValidateNonNegative(gutterWidthTwips, nameof(gutterWidthTwips));
        GutterWidthTwips = gutterWidthTwips;
        RtlGutter = rtlGutter;
        return this;
    }

    /// <summary>Sets header and footer distances from the page edges in twips.</summary>
    public RtfPageSetup SetHeaderFooterDistance(int? headerDistanceTwips = null, int? footerDistanceTwips = null) {
        ValidateNonNegative(headerDistanceTwips, nameof(headerDistanceTwips));
        ValidateNonNegative(footerDistanceTwips, nameof(footerDistanceTwips));
        HeaderDistanceTwips = headerDistanceTwips;
        FooterDistanceTwips = footerDistanceTwips;
        return this;
    }

    /// <summary>Sets page numbering controls.</summary>
    public RtfPageSetup SetPageNumbering(
        int? start = null,
        bool? restart = null,
        RtfPageNumberFormat? format = null,
        int? positionXTwips = null,
        int? positionYTwips = null) {
        ValidatePositive(start, nameof(start));
        ValidateNonNegative(positionXTwips, nameof(positionXTwips));
        ValidateNonNegative(positionYTwips, nameof(positionYTwips));
        PageNumberStart = start;
        PageNumberRestart = restart;
        PageNumberFormat = format;
        PageNumberPositionXTwips = positionXTwips;
        PageNumberPositionYTwips = positionYTwips;
        return this;
    }

    /// <summary>Sets document margins in twips.</summary>
    public RtfPageSetup SetMargins(int? leftTwips = null, int? rightTwips = null, int? topTwips = null, int? bottomTwips = null) {
        ValidateNonNegative(leftTwips, nameof(leftTwips));
        ValidateNonNegative(rightTwips, nameof(rightTwips));
        ValidateNonNegative(topTwips, nameof(topTwips));
        ValidateNonNegative(bottomTwips, nameof(bottomTwips));
        MarginLeftTwips = leftTwips;
        MarginRightTwips = rightTwips;
        MarginTopTwips = topTwips;
        MarginBottomTwips = bottomTwips;
        return this;
    }

    /// <summary>Sets whether landscape orientation should be emitted.</summary>
    public RtfPageSetup SetLandscape(bool landscape = true) {
        Landscape = landscape;
        return this;
    }

    /// <summary>Sets whether first-page header/footer variants should be enabled.</summary>
    public RtfPageSetup SetDifferentFirstPageHeaderFooter(bool enabled = true) {
        DifferentFirstPageHeaderFooter = enabled;
        return this;
    }

    internal bool HasAnyValue =>
        PaperWidthTwips.HasValue ||
        PaperHeightTwips.HasValue ||
        PrinterPaperSize.HasValue ||
        FirstPagePaperSource.HasValue ||
        OtherPagesPaperSource.HasValue ||
        MarginLeftTwips.HasValue ||
        MarginRightTwips.HasValue ||
        MarginTopTwips.HasValue ||
        MarginBottomTwips.HasValue ||
        GutterWidthTwips.HasValue ||
        HeaderDistanceTwips.HasValue ||
        FooterDistanceTwips.HasValue ||
        PageNumberStart.HasValue ||
        PageNumberRestart.HasValue ||
        PageNumberPositionXTwips.HasValue ||
        PageNumberPositionYTwips.HasValue ||
        PageNumberFormat.HasValue ||
        PageBorders.HasAnyValue ||
        Landscape ||
        DifferentFirstPageHeaderFooter ||
        RtlGutter;

    internal void Clear() {
        PaperWidthTwips = null;
        PaperHeightTwips = null;
        PrinterPaperSize = null;
        FirstPagePaperSource = null;
        OtherPagesPaperSource = null;
        MarginLeftTwips = null;
        MarginRightTwips = null;
        MarginTopTwips = null;
        MarginBottomTwips = null;
        GutterWidthTwips = null;
        HeaderDistanceTwips = null;
        FooterDistanceTwips = null;
        PageNumberStart = null;
        PageNumberRestart = null;
        PageNumberPositionXTwips = null;
        PageNumberPositionYTwips = null;
        PageNumberFormat = null;
        PageBorders.Clear();
        Landscape = false;
        DifferentFirstPageHeaderFooter = false;
        RtlGutter = false;
    }

    private static void ValidateNonNegative(int? value, string parameterName) {
        if (value.HasValue && value.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Margin cannot be negative.");
        }
    }

    private static void ValidatePositive(int? value, string parameterName) {
        if (value.HasValue && value.Value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Page number start must be greater than zero.");
        }
    }
}
