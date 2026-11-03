namespace OfficeIMO.Pdf;

/// <summary>
/// Page size, margins, orientation and running headers and footers of a PDF document, composed page or section.
/// </summary>
/// <remarks>
/// The setup is a view over its <see cref="PdfCreateOptions"/>; it holds no state of its own, so cloned and
/// composed-page options expose their own snapshot through their own setup.
/// </remarks>
public sealed class PdfPageSetup {
    private readonly PdfCreateOptions _options;

    internal PdfPageSetup(PdfCreateOptions options) {
        _options = options;
    }

    /// <summary>Gets or sets the page size. Default is Letter (612 by 792 points).</summary>
    public PdfPageSize Size {
        get => _options.PageSize;
        set => _options.PageSize = value;
    }

    /// <summary>Gets or sets the page orientation. Setting it swaps the page width and height when needed.</summary>
    public OfficePageOrientation Orientation {
        get => _options.PageWidth > _options.PageHeight ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
        set => _options.PageSize = _options.PageSize.WithOrientation(value);
    }

    /// <summary>Gets the page margins. Sides that were not set follow the document theme.</summary>
    public PdfPageSetupMargins Margins => new PdfPageSetupMargins(_options);

    /// <summary>Swaps the left and right margins on even visible page numbers. Numbering starts and restarts determine parity.</summary>
    public bool MirrorMargins {
        get => _options.MirrorMargins;
        set => _options.MirrorMargins = value;
    }

    /// <summary>When true, page 1 of each section uses <see cref="FirstPageHeader"/> and <see cref="FirstPageFooter"/>.</summary>
    public bool DifferentFirstPage {
        get => _options.DifferentFirstPageHeaderFooter;
        set => _options.DifferentFirstPageHeaderFooter = value;
    }

    /// <summary>When true, even pages use <see cref="EvenPageHeader"/> and <see cref="EvenPageFooter"/>.</summary>
    public bool DifferentOddAndEvenPages {
        get => _options.DifferentOddAndEvenPagesHeaderFooter;
        set => _options.DifferentOddAndEvenPagesHeaderFooter = value;
    }

    /// <summary>
    /// Uses visible page-number parity for odd/even headers, footers and watermarks. First-page variants still refer to the
    /// first physical page of each section. The default uses section-relative page parity.
    /// </summary>
    public bool UseVisiblePageNumberParity {
        get => _options.UsePageNumberParityForHeaderFooter;
        set => _options.UsePageNumberParityForHeaderFooter = value;
    }

    /// <summary>Gets the running header.</summary>
    public PdfHeaderFooter Header => new PdfHeaderFooter(_options, isHeader: true, PdfRunningContentVariant.Default);
    /// <summary>Gets the page-1 header used when <see cref="DifferentFirstPage"/> is true.</summary>
    public PdfHeaderFooter FirstPageHeader => new PdfHeaderFooter(_options, isHeader: true, PdfRunningContentVariant.FirstPage);
    /// <summary>Gets the even-page header used when <see cref="DifferentOddAndEvenPages"/> is true.</summary>
    public PdfHeaderFooter EvenPageHeader => new PdfHeaderFooter(_options, isHeader: true, PdfRunningContentVariant.EvenPages);
    /// <summary>Gets the running footer.</summary>
    public PdfHeaderFooter Footer => new PdfHeaderFooter(_options, isHeader: false, PdfRunningContentVariant.Default);
    /// <summary>Gets the page-1 footer used when <see cref="DifferentFirstPage"/> is true.</summary>
    public PdfHeaderFooter FirstPageFooter => new PdfHeaderFooter(_options, isHeader: false, PdfRunningContentVariant.FirstPage);
    /// <summary>Gets the even-page footer used when <see cref="DifferentOddAndEvenPages"/> is true.</summary>
    public PdfHeaderFooter EvenPageFooter => new PdfHeaderFooter(_options, isHeader: false, PdfRunningContentVariant.EvenPages);
}

/// <summary>
/// Page margins of a <see cref="PdfPageSetup"/>. Each side either follows the document theme or holds an explicit value.
/// </summary>
public sealed class PdfPageSetupMargins {
    private readonly PdfCreateOptions _options;

    internal PdfPageSetupMargins(PdfCreateOptions options) {
        _options = options;
    }

    /// <summary>Left margin. Default comes from the theme (72 points without one).</summary>
    public OfficeLength Left { get => _options.MarginLeft; set => _options.MarginLeft = value; }
    /// <summary>Top margin. Default comes from the theme (72 points without one).</summary>
    public OfficeLength Top { get => _options.MarginTop; set => _options.MarginTop = value; }
    /// <summary>Right margin. Default comes from the theme (72 points without one).</summary>
    public OfficeLength Right { get => _options.MarginRight; set => _options.MarginRight = value; }
    /// <summary>Bottom margin. Default comes from the theme (72 points without one).</summary>
    public OfficeLength Bottom { get => _options.MarginBottom; set => _options.MarginBottom = value; }

    /// <summary>Sets all four sides explicitly from a reusable margin value.</summary>
    public void Set(PdfPageMargins margins) => _options.Margins = margins;

    /// <summary>Returns the current effective margins as a reusable value.</summary>
    public PdfPageMargins ToPdfPageMargins() => _options.Margins;
}

/// <summary>
/// One running header or footer slot of a <see cref="PdfPageSetup"/>: its text format, styled segments, three-zone text,
/// images and shapes, plus the text style shared by every header (or every footer) variant.
/// </summary>
/// <remarks>
/// <see cref="StandardFont"/>, <see cref="FontFamily"/>, <see cref="FontSize"/>, <see cref="Color"/>,
/// <see cref="Alignment"/> and <see cref="Offset"/> belong to the header (or footer) as a whole: the first-page and
/// even-page variants read and write the same values as the running header (or footer).
/// </remarks>
public sealed class PdfHeaderFooter {
    private readonly PdfCreateOptions _options;
    private readonly bool _isHeader;
    private readonly PdfRunningContentVariant _variant;

    internal PdfHeaderFooter(PdfCreateOptions options, bool isHeader, PdfRunningContentVariant variant) {
        _options = options;
        _isHeader = isHeader;
        _variant = variant;
    }

    /// <summary>
    /// When true, renders <see cref="Format"/>. The running header and footer start hidden; the first-page and even-page
    /// variants start shown and render whenever their format is not empty.
    /// </summary>
    public bool Show {
        get => _options.GetRunningTextShown(_isHeader, _variant);
        set => _options.SetRunningTextShown(_isHeader, _variant, value);
    }

    /// <summary>
    /// Text format with <c>{page}</c> and <c>{pages}</c> tokens. The running footer defaults to
    /// <c>"Page {page} of {pages}"</c>; other slots default to empty.
    /// </summary>
    public string Format {
        get => _options.GetRunningTextFormat(_isHeader, _variant);
        set => _options.SetRunningTextFormat(_isHeader, _variant, value);
    }

    /// <summary>Styled text and page-token segments. When set, they replace <see cref="Format"/>.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfFooterSegment>? Segments {
        get => _options.GetRunningTextSegments(_isHeader, _variant);
        set => _options.SetRunningTextSegments(_isHeader, _variant, value);
    }

    /// <summary>Header or footer standard font.</summary>
    public PdfStandardFont StandardFont {
        get => _isHeader ? _options.HeaderFont : _options.FooterFont;
        set {
            if (_isHeader) _options.HeaderFont = value;
            else _options.FooterFont = value;
        }
    }

    internal bool HasExplicitStandardFont => _isHeader ? _options.HasExplicitHeaderFont : _options.HasExplicitFooterFont;

    /// <summary>Optional registered named font family. When null, <see cref="StandardFont"/> is used.</summary>
    public string? FontFamily {
        get => _isHeader ? _options.HeaderFontFamily : _options.FooterFontFamily;
        set {
            if (_isHeader) _options.HeaderFontFamily = value;
            else _options.FooterFontFamily = value;
        }
    }

    /// <summary>Font size. Headers default to 9 points; footers follow the theme footer size.</summary>
    public OfficeLength FontSize {
        get => _isHeader ? _options.HeaderFontSize : _options.FooterFontSize;
        set {
            if (_isHeader) _options.HeaderFontSize = value;
            else _options.FooterFontSize = value;
        }
    }

    /// <summary>Text color. When null, headers keep the current PDF fill color and footers follow the theme footer color.</summary>
    public global::OfficeIMO.OfficeColor? Color {
        get => _isHeader ? _options.HeaderTextColor : _options.FooterTextColor;
        set {
            if (_isHeader) _options.HeaderTextColor = value;
            else _options.FooterTextColor = value;
        }
    }

    /// <summary>Horizontal alignment: left, center (default) or right.</summary>
    public OfficeAlignment Alignment {
        get => _isHeader ? _options.HeaderAlign : _options.FooterAlign;
        set {
            if (_isHeader) _options.HeaderAlign = value;
            else _options.FooterAlign = value;
        }
    }

    /// <summary>Baseline distance from the margin: above the top margin for headers, below the bottom margin for footers. Default 18 points.</summary>
    public OfficeLength Offset {
        get => _isHeader ? _options.HeaderOffsetY : _options.FooterOffsetY;
        set {
            if (_isHeader) _options.HeaderOffsetY = value;
            else _options.FooterOffsetY = value;
        }
    }

    /// <summary>
    /// Renders <paramref name="format"/> in this slot, replacing segments and zones. Shows the running slot, or turns on the
    /// matching <see cref="PdfPageSetup"/> variant flag for a first-page or even-page slot.
    /// </summary>
    public void SetText(string format) {
        Guard.NotNull(format, nameof(format));
        _options.SetRunningText(_isHeader, _variant, format);
    }

    /// <summary>Renders styled segments in this slot, replacing its format text and zones.</summary>
    public void SetText(System.Collections.Generic.IEnumerable<PdfFooterSegment> segments) {
        Guard.NotNull(segments, nameof(segments));
        _options.SetRunningText(_isHeader, _variant, segments);
    }

    /// <summary>Renders left, center and right text zones on one line. Supports <c>{page}</c> and <c>{pages}</c>.</summary>
    public void SetZones(string? left, string? center, string? right) =>
        _options.SetRunningZones(_isHeader, _variant, left, center, right);

    /// <summary>Renders independently styled left, center and right zones on one line.</summary>
    public void SetStyledZones(
        System.Collections.Generic.IEnumerable<PdfFooterSegment>? left,
        System.Collections.Generic.IEnumerable<PdfFooterSegment>? center,
        System.Collections.Generic.IEnumerable<PdfFooterSegment>? right) =>
        _options.SetRunningZoneSegments(_isHeader, _variant, ToList(left), ToList(center), ToList(right));

    /// <summary>Adds an image to this slot.</summary>
    public void AddImage(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        _options.AddRunningImage(_isHeader, _variant, image);
    }

    /// <summary>Adds a shape to this slot.</summary>
    public void AddShape(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        _options.AddRunningShape(_isHeader, _variant, shape);
    }

    private static System.Collections.Generic.List<PdfFooterSegment>? ToList(System.Collections.Generic.IEnumerable<PdfFooterSegment>? segments) =>
        segments == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(segments);
}
