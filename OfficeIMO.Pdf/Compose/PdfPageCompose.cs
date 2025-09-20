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
    public PdfPageCompose Size(PageSize size) { Options.PageWidth = size.Width; Options.PageHeight = size.Height; return this; }
    /// <summary>Sets custom page size in points.</summary>
    public PdfPageCompose Size(double width, double height) { Options.PageWidth = width; Options.PageHeight = height; return this; }
    /// <summary>Sets uniform page margins (all sides in points).</summary>
    public PdfPageCompose Margin(double all) { Options.MarginLeft = Options.MarginRight = Options.MarginTop = Options.MarginBottom = all; return this; }
    /// <summary>Sets page margins (left, top, right, bottom in points).</summary>
    public PdfPageCompose Margin(double left, double top, double right, double bottom) { Options.MarginLeft = left; Options.MarginTop = top; Options.MarginRight = right; Options.MarginBottom = bottom; return this; }

    /// <summary>Configures default text style for the page.</summary>
    public PdfPageCompose DefaultTextStyle(System.Action<PdfTextStyleCompose> style) { var s = new PdfTextStyleCompose(Options); style(s); return this; }
    /// <summary>Builds the page content using a column/row flow.</summary>
    public PdfPageCompose Content(System.Action<PdfContentCompose> build) { var c = new PdfContentCompose(_doc); build(c); return this; }
    /// <summary>Defines the footer layout and content.</summary>
    public PdfPageCompose Footer(System.Action<PdfFooterCompose> build) { var f = new PdfFooterCompose(Options); build(f); return this; }
}

