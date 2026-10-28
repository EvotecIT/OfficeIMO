namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private double _pageWidth = 612;
    private double _pageHeight = 792;
    private long _pageSizeConfigurationVersion;
    private double? _marginLeft;
    private double? _marginRight;
    private double? _marginTop;
    private double? _marginBottom;

    internal long PageSizeConfigurationState => _pageSizeConfigurationVersion;

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

    /// <summary>Page width in points (1 pt = 1/72 in). Default is 612 (Letter 8.5in).</summary>
    public double PageWidth {
        get => _pageWidth;
        set {
            _pageWidth = value;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page height in points. Default is 792 (Letter 11in).</summary>
    public double PageHeight {
        get => _pageHeight;
        set {
            _pageHeight = value;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page size in points.</summary>
    public PdfPageSize PageSize {
        get => new PdfPageSize(PageWidth, PageHeight);
        set {
            Guard.Positive(value.Width, nameof(PageSize));
            Guard.Positive(value.Height, nameof(PageSize));
            _pageWidth = value.Width;
            _pageHeight = value.Height;
            _pageSizeConfigurationVersion++;
        }
    }
    /// <summary>Page orientation inferred from the current page size.</summary>
    public OfficePageOrientation PageOrientation => PageWidth > PageHeight ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
    /// <summary>Left margin in points. Defaults to the selected theme's page margin.</summary>
    public double MarginLeft { get => _marginLeft ?? EffectiveTheme.Page.MarginLeft.ToPoints(); set => _marginLeft = value; }
    /// <summary>Right margin in points. Defaults to the selected theme's page margin.</summary>
    public double MarginRight { get => _marginRight ?? EffectiveTheme.Page.MarginRight.ToPoints(); set => _marginRight = value; }
    /// <summary>Swaps the left and right margins on even visible page numbers. Numbering starts and restarts determine parity.</summary>
    public bool MirrorMargins { get; set; }
    /// <summary>Top margin in points. Defaults to the selected theme's page margin.</summary>
    public double MarginTop { get => _marginTop ?? EffectiveTheme.Page.MarginTop.ToPoints(); set => _marginTop = value; }
    /// <summary>Bottom margin in points. Defaults to the selected theme's page margin.</summary>
    public double MarginBottom { get => _marginBottom ?? EffectiveTheme.Page.MarginBottom.ToPoints(); set => _marginBottom = value; }
    /// <summary>Page margins in points.</summary>
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
