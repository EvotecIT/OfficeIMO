namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private double _pageWidth = 612;
    private double _pageHeight = 792;
    private long _pageSizeConfigurationVersion;

    internal long PageSizeConfigurationState => _pageSizeConfigurationVersion;

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
    public PageSize PageSize {
        get => new PageSize(PageWidth, PageHeight);
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
    /// <summary>Left margin in points. Default 72 (1 inch).</summary>
    public double MarginLeft { get; set; } = 72; // 1 in
    /// <summary>Right margin in points. Default 72 (1 inch).</summary>
    public double MarginRight { get; set; } = 72;
    /// <summary>Top margin in points. Default 72 (1 inch).</summary>
    public double MarginTop { get; set; } = 72;
    /// <summary>Bottom margin in points. Default 72 (1 inch).</summary>
    public double MarginBottom { get; set; } = 72;
    /// <summary>Page margins in points.</summary>
    public PageMargins Margins {
        get => new PageMargins(MarginLeft, MarginTop, MarginRight, MarginBottom);
        set {
            MarginLeft = value.Left;
            MarginTop = value.Top;
            MarginRight = value.Right;
            MarginBottom = value.Bottom;
        }
    }
}
