namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private OfficeLength? _marginLeft;
    private OfficeLength? _marginRight;
    private OfficeLength? _marginTop;
    private OfficeLength? _marginBottom;
    private OfficeLength _pageWidth = OfficeLength.Points(612);
    private OfficeLength _pageHeight = OfficeLength.Points(792);
    private long _pageSizeConfigurationVersion;

    internal long PageSizeConfigurationState => _pageSizeConfigurationVersion;

    /// <summary>True when any page margin was set explicitly instead of coming from the theme.</summary>
    internal bool HasExplicitMargins => _marginLeft.HasValue || _marginRight.HasValue || _marginTop.HasValue || _marginBottom.HasValue;

    private PdfPageParity? _pageStartParity;
    /// <summary>Uses the page number continuing the preceding group for start parity. Numbering restarts are applied after padding. Default is false, which uses the physical page index.</summary>
    public bool UseContinuingPageNumberForStartParity { get; set; }
    /// <summary>Optional start parity for a composed page or section. Uses the physical page index unless <see cref="UseContinuingPageNumberForStartParity"/> is enabled. A blank page is inserted before a later section when needed; the first section starts without padding.</summary>
    public PdfPageParity? PageStartParity {
        get => _pageStartParity;
        set {
            if (value.HasValue && value != PdfPageParity.Odd && value != PdfPageParity.Even) {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _pageStartParity = value;
        }
    }

    /// <summary>Page width. Default is 612 points (Letter 8.5in).</summary>
    public OfficeLength PageWidth {
        get => _pageWidth;
        set {
            _pageWidth = value;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page height. Default is 792 points (Letter 11in).</summary>
    public OfficeLength PageHeight {
        get => _pageHeight;
        set {
            _pageHeight = value;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page size.</summary>
    public PdfPageSize PageSize {
        get => new PdfPageSize(PageWidth, PageHeight);
        set {
            Guard.Positive(value.Width.ToPoints(), nameof(PageSize));
            Guard.Positive(value.Height.ToPoints(), nameof(PageSize));
            _pageWidth = value.Width;
            _pageHeight = value.Height;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page orientation inferred from the current page size.</summary>
    public OfficePageOrientation PageOrientation => PageWidth > PageHeight ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
    /// <summary>Left margin. Default 72 points (1 inch).</summary>
    public OfficeLength MarginLeft { get => _marginLeft ?? (_useOfficeThemeDefaults ? EffectiveTheme.Page.MarginLeft : OfficeLength.Points(72)); set => _marginLeft = value; }
    /// <summary>Right margin. Default 72 points (1 inch).</summary>
    public OfficeLength MarginRight { get => _marginRight ?? (_useOfficeThemeDefaults ? EffectiveTheme.Page.MarginRight : OfficeLength.Points(72)); set => _marginRight = value; }
    /// <summary>Swaps the left and right margins on even visible page numbers. Numbering starts and restarts determine parity.</summary>
    public bool MirrorMargins { get; set; }
    /// <summary>Top margin. Default 72 points (1 inch).</summary>
    public OfficeLength MarginTop { get => _marginTop ?? (_useOfficeThemeDefaults ? EffectiveTheme.Page.MarginTop : OfficeLength.Points(72)); set => _marginTop = value; }
    /// <summary>Bottom margin. Default 72 points (1 inch).</summary>
    public OfficeLength MarginBottom { get => _marginBottom ?? (_useOfficeThemeDefaults ? EffectiveTheme.Page.MarginBottom : OfficeLength.Points(72)); set => _marginBottom = value; }
    /// <summary>Page margins.</summary>
    public PdfPageMargins Margins {
        get => new PdfPageMargins(MarginLeft, MarginTop, MarginRight, MarginBottom);
        set {
            MarginLeft = value.Left;
            MarginTop = value.Top;
            MarginRight = value.Right;
            MarginBottom = value.Bottom;
        }
    }
}
