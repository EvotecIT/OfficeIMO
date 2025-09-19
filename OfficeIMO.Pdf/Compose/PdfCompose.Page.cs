namespace OfficeIMO.Pdf;

public sealed class PdfCompose {
    private readonly PdfDoc _doc;
    internal PdfCompose(PdfDoc doc) { _doc = doc; }
    public PdfCompose Page(System.Action<PdfPageCompose> configure) { var p = new PdfPageCompose(_doc); configure(p); return this; }
}

public sealed class PdfPageCompose {
    private readonly PdfDoc _doc;
    internal PdfOptions Options => _doc.Options;
    internal PdfPageCompose(PdfDoc doc) { _doc = doc; }

    public PdfPageCompose Size(PageSize size) { Options.PageWidth = size.Width; Options.PageHeight = size.Height; return this; }
    public PdfPageCompose Size(double width, double height) { Options.PageWidth = width; Options.PageHeight = height; return this; }
    public PdfPageCompose Margin(double all) { Options.MarginLeft = Options.MarginRight = Options.MarginTop = Options.MarginBottom = all; return this; }
    public PdfPageCompose Margin(double left, double top, double right, double bottom) { Options.MarginLeft = left; Options.MarginTop = top; Options.MarginRight = right; Options.MarginBottom = bottom; return this; }

    public PdfPageCompose DefaultTextStyle(System.Action<PdfTextStyleCompose> style) { var s = new PdfTextStyleCompose(Options); style(s); return this; }

    public PdfPageCompose Content(System.Action<PdfContentCompose> build) { var c = new PdfContentCompose(_doc); build(c); return this; }
    public PdfPageCompose Footer(System.Action<PdfFooterCompose> build) { var f = new PdfFooterCompose(Options); build(f); return this; }
}

public readonly struct PageSize {
    public double Width { get; }
    public double Height { get; }
    public PageSize(double width, double height) { Width = width; Height = height; }
}

public static class PageSizes {
    public static PageSize A5 => new PageSize(420, 595);
    public static PageSize A4 => new PageSize(595, 842);
    public static PageSize Letter => new PageSize(612, 792);
    public static PageSize Legal => new PageSize(612, 1008);
}

public sealed class PdfTextStyleCompose {
    private readonly PdfOptions _opts;
    internal PdfTextStyleCompose(PdfOptions opts) { _opts = opts; }
    public PdfTextStyleCompose FontSize(double size) { _opts.DefaultFontSize = size; return this; }
    public PdfTextStyleCompose Color(PdfColor color) { _opts.DefaultTextColor = color; return this; }
    public PdfTextStyleCompose Font(PdfStandardFont font) { _opts.DefaultFont = font; return this; }
}

