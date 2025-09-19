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

public sealed class PdfContentCompose {
    private readonly PdfDoc _doc;
    internal PdfContentCompose(PdfDoc doc) { _doc = doc; }
    public PdfContentCompose PaddingBottom(double points) { /* reserved for future */ return this; }
    public PdfContentCompose Column(System.Action<PdfColumnCompose> build) { var col = new PdfColumnCompose(_doc); build(col); return this; }
    public PdfContentCompose Row(System.Action<PdfRowCompose> build) { var row = new PdfRowCompose(_doc); build(row); row.Commit(); return this; }
}

public sealed class PdfColumnCompose {
    private readonly PdfDoc _doc;
    internal PdfColumnCompose(PdfDoc doc) { _doc = doc; }
    public PdfItemCompose Item() => new PdfItemCompose(_doc);
}

public sealed class PdfItemCompose {
    private readonly PdfDoc _doc;
    internal PdfItemCompose(PdfDoc doc) { _doc = doc; }
    public PdfItemCompose PageBreak() { _doc.PageBreak(); return this; }
    public PdfItemCompose H1(string text) { _doc.H1(text); return this; }
    public PdfItemCompose H2(string text) { _doc.H2(text); return this; }
    public PdfItemCompose H3(string text) { _doc.H3(text); return this; }
    public PdfItemCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    public PdfItemCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    public PdfItemCompose Element(System.Action<PdfElementCompose> build) { var el = new PdfElementCompose(_doc); build(el); return this; }
    public PdfItemCompose HR(double thickness = 0.5, PdfColor? color = null, double spacingBefore = 6, double spacingAfter = 6) { _doc.HR(thickness, color, spacingBefore, spacingAfter); return this; }
    public PdfItemCompose PanelParagraph(System.Action<PdfParagraphBuilder> build, PanelStyle style, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.PanelParagraph(build, style, align, defaultColor); return this; }
    public PdfItemCompose Image(byte[] jpegBytes, double width, double height, PdfAlign align = PdfAlign.Left) { _doc.Image(jpegBytes, width, height, align); return this; }
}

public sealed class PdfRowCompose {
    private readonly PdfDoc _doc;
    private readonly RowBlock _row = new RowBlock();
    internal PdfRowCompose(PdfDoc doc) { _doc = doc; }
    public PdfRowCompose Column(double widthPercent, System.Action<PdfRowColumnCompose> build) {
        var col = new RowColumn(widthPercent);
        var cc = new PdfRowColumnCompose(col);
        build(cc);
        _row.Columns.Add(col);
        return this;
    }
    internal void Commit() { _doc.AddRow(_row); }
}

public sealed class PdfRowColumnCompose {
    private readonly RowColumn _col;
    internal PdfRowColumnCompose(RowColumn col) { _col = col; }
    public PdfRowColumnCompose H1(string text) { _col.Blocks.Add(new HeadingBlock(1, text, PdfAlign.Left, null)); return this; }
    public PdfRowColumnCompose H2(string text) { _col.Blocks.Add(new HeadingBlock(2, text, PdfAlign.Left, null)); return this; }
    public PdfRowColumnCompose H3(string text) { _col.Blocks.Add(new HeadingBlock(3, text, PdfAlign.Left, null)); return this; }
    public PdfRowColumnCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        var b = new PdfParagraphBuilder(align, defaultColor);
        build(b);
        _col.Blocks.Add(new RichParagraphBlock(b.Build().Runs, align, defaultColor));
        return this;
    }
    public PdfRowColumnCompose HR(double thickness = 0.5, PdfColor? color = null, double spacingBefore = 6, double spacingAfter = 6) { _col.Blocks.Add(new HorizontalRuleBlock(thickness, color ?? PdfColor.Gray, spacingBefore, spacingAfter)); return this; }
    public PdfRowColumnCompose Image(byte[] jpegBytes, double width, double height, PdfAlign align = PdfAlign.Left) { _col.Blocks.Add(new ImageBlock(jpegBytes, width, height, align)); return this; }
}

public sealed class PdfElementCompose {
    private readonly PdfDoc _doc;
    internal PdfElementCompose(PdfDoc doc) { _doc = doc; }
    public PdfElementCompose H1(string text) { _doc.H1(text); return this; }
    public PdfElementCompose H2(string text) { _doc.H2(text); return this; }
    public PdfElementCompose H3(string text) { _doc.H3(text); return this; }
    public PdfElementCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
}

public sealed class PdfFooterCompose {
    private readonly PdfOptions _opts;
    internal PdfFooterCompose(PdfOptions opts) { _opts = opts; }
    public PdfFooterCompose AlignLeft() { _opts.FooterAlign = PdfAlign.Left; return this; }
    public PdfFooterCompose AlignCenter() { _opts.FooterAlign = PdfAlign.Center; return this; }
    public PdfFooterCompose AlignRight() { _opts.FooterAlign = PdfAlign.Right; return this; }
    public PdfFooterCompose PageNumber() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}"; return this; }
    public PdfFooterCompose PageNumberWithTotal() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}/{pages}"; return this; }
    public PdfFooterCompose Text(System.Action<FooterTextBuilder> build) {
        _opts.FooterSegments = new System.Collections.Generic.List<FooterSegment>();
        var b = new FooterTextBuilder(_opts.FooterSegments);
        build(b);
        _opts.ShowPageNumbers = true; // might be needed when builder inserts tokens
        return this;
    }
}

public sealed class FooterTextBuilder {
    private readonly System.Collections.Generic.List<FooterSegment> _segments;
    internal FooterTextBuilder(System.Collections.Generic.List<FooterSegment> segs) { _segments = segs; }
    public FooterTextBuilder Text(string s) { _segments.Add(new FooterSegment(FooterSegmentKind.Text, s)); return this; }
    public FooterTextBuilder CurrentPage() { _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber)); return this; }
    public FooterTextBuilder TotalPages() { _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages)); return this; }
}
