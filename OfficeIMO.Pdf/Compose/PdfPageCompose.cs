namespace OfficeIMO.Pdf;

/// <summary>
/// Page-level configuration (size, margins, default styles) and content/footers.
/// </summary>
public class PdfPageCompose {
    private readonly PdfDoc _doc;
    private readonly PdfOptions _options;
    internal PdfOptions Options => _options;
    internal PdfPageCompose(PdfDoc doc, PdfOptions options) { _doc = doc; _options = options; }

    /// <summary>Sets page size using a predefined <see cref="PageSize"/>.</summary>
    public PdfPageCompose Size(PageSize size) {
        Guard.Positive(size.Width, nameof(size));
        Guard.Positive(size.Height, nameof(size));
        Options.PageSize = size;
        return this;
    }
    /// <summary>Sets custom page size in points.</summary>
    public PdfPageCompose Size(double width, double height) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Options.PageWidth = width;
        Options.PageHeight = height;
        return this;
    }
    /// <summary>Sets page orientation while preserving the current page size dimensions.</summary>
    public PdfPageCompose Orientation(PdfPageOrientation orientation) {
        var oriented = new PageSize(Options.PageWidth, Options.PageHeight).WithOrientation(orientation);
        Options.PageWidth = oriented.Width;
        Options.PageHeight = oriented.Height;
        return this;
    }
    /// <summary>Sets or clears the page background color.</summary>
    public PdfPageCompose Background(PdfColor? color) {
        Options.BackgroundColor = color;
        return this;
    }
    /// <summary>Sets page orientation to portrait while preserving the current page size dimensions.</summary>
    public PdfPageCompose Portrait() => Orientation(PdfPageOrientation.Portrait);
    /// <summary>Sets page orientation to landscape while preserving the current page size dimensions.</summary>
    public PdfPageCompose Landscape() => Orientation(PdfPageOrientation.Landscape);
    /// <summary>Sets uniform page margins (all sides in points).</summary>
    public PdfPageCompose Margin(double all) {
        Guard.NonNegative(all, nameof(all));
        Options.MarginLeft = Options.MarginRight = Options.MarginTop = Options.MarginBottom = all;
        return this;
    }
    /// <summary>Sets page margins from a reusable margin value.</summary>
    public PdfPageCompose Margin(PageMargins margins) {
        Options.Margins = margins;
        return this;
    }
    /// <summary>Sets page margins (left, top, right, bottom in points).</summary>
    public PdfPageCompose Margin(double left, double top, double right, double bottom) {
        Guard.NonNegative(left, nameof(left));
        Guard.NonNegative(top, nameof(top));
        Guard.NonNegative(right, nameof(right));
        Guard.NonNegative(bottom, nameof(bottom));
        Options.MarginLeft = left;
        Options.MarginTop = top;
        Options.MarginRight = right;
        Options.MarginBottom = bottom;
        return this;
    }

    /// <summary>Sets the first visible page number for this page or section flow.</summary>
    public PdfPageCompose PageNumberStart(int start) {
        Options.PageNumberStart = start;
        return this;
    }

    /// <summary>Sets the visible page-number style for this page or section flow.</summary>
    public PdfPageCompose PageNumberStyle(PdfPageNumberStyle style) {
        Options.PageNumberStyle = style;
        return this;
    }

    /// <summary>Applies reusable page-scoped default styles.</summary>
    public PdfPageCompose Theme(PdfTheme theme) { Guard.NotNull(theme, nameof(theme)); theme.Clone().ApplyTo(Options); return this; }
    /// <summary>Configures default text style for the page.</summary>
    public PdfPageCompose DefaultTextStyle(System.Action<PdfTextStyleCompose> style) { Guard.NotNull(style, nameof(style)); var s = new PdfTextStyleCompose(Options); style(s); return this; }
    /// <summary>Configures default text style for the page from a reusable text style object.</summary>
    public PdfPageCompose DefaultTextStyle(PdfTextStyle style) { Guard.NotNull(style, nameof(style)); style.Clone().ApplyTo(Options); return this; }
    /// <summary>Configures the default paragraph style for page paragraphs that do not provide an explicit style.</summary>
    public PdfPageCompose DefaultParagraphStyle(PdfParagraphStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultParagraphStyle = style; return this; }
    /// <summary>Configures the default table style for page tables that do not provide an explicit style.</summary>
    public PdfPageCompose DefaultTableStyle(PdfTableStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultTableStyle = style; return this; }
    /// <summary>Configures the default table style for the page from a supported Word table style name.</summary>
    public PdfPageCompose DefaultTableStyle(string wordTableStyleName) { Options.DefaultTableStyle = TableStyles.FromWordTableStyle(wordTableStyleName); return this; }
    /// <summary>Configures the default style for a built-in heading level on the page.</summary>
    public PdfPageCompose DefaultHeadingStyle(int level, PdfHeadingStyle style) { Guard.NotNull(style, nameof(style)); Options.SetDefaultHeadingStyle(level, style); return this; }
    /// <summary>Configures the default style for page bullet and numbered lists.</summary>
    public PdfPageCompose DefaultListStyle(PdfListStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultListStyle = style; return this; }
    /// <summary>Configures the default style for page panel paragraphs.</summary>
    public PdfPageCompose DefaultPanelStyle(PanelStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultPanelStyle = style; return this; }
    /// <summary>Configures the default style for page horizontal rules.</summary>
    public PdfPageCompose DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultHorizontalRuleStyle = style; return this; }
    /// <summary>Configures the default style for page images.</summary>
    public PdfPageCompose DefaultImageStyle(PdfImageStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultImageStyle = style; return this; }
    /// <summary>Configures the default placement style for page drawing objects.</summary>
    public PdfPageCompose DefaultDrawingStyle(PdfDrawingStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultDrawingStyle = style; return this; }
    /// <summary>Configures the default row/column layout style for the page.</summary>
    public PdfPageCompose DefaultRowStyle(PdfRowStyle style) { Guard.NotNull(style, nameof(style)); Options.DefaultRowStyle = style; return this; }
    /// <summary>Builds the page content using a column/row flow.</summary>
    public PdfPageCompose Content(System.Action<PdfContentCompose> build) { Guard.NotNull(build, nameof(build)); var c = new PdfContentCompose(_doc); build(c); return this; }
    /// <summary>Defines the header layout and content.</summary>
    public PdfPageCompose Header(System.Action<PdfHeaderCompose> build) { Guard.NotNull(build, nameof(build)); var h = new PdfHeaderCompose(Options); build(h); return this; }
    /// <summary>Defines the footer layout and content.</summary>
    public PdfPageCompose Footer(System.Action<PdfFooterCompose> build) { Guard.NotNull(build, nameof(build)); var f = new PdfFooterCompose(Options); build(f); return this; }
}

