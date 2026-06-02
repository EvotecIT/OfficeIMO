namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>Page width in points (1 pt = 1/72 in). Default is 612 (Letter 8.5in).</summary>
    public double PageWidth { get; set; } = 612; // Letter 8.5in * 72
    /// <summary>Page height in points. Default is 792 (Letter 11in).</summary>
    public double PageHeight { get; set; } = 792; // Letter 11in * 72
    /// <summary>Page size in points.</summary>
    public PageSize PageSize {
        get => new PageSize(PageWidth, PageHeight);
        set {
            Guard.Positive(value.Width, nameof(PageSize));
            Guard.Positive(value.Height, nameof(PageSize));
            PageWidth = value.Width;
            PageHeight = value.Height;
        }
    }
    /// <summary>Page orientation inferred from the current page size.</summary>
    public PdfPageOrientation PageOrientation => PageWidth > PageHeight ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
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
