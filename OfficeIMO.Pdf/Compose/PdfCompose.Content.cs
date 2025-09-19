namespace OfficeIMO.Pdf;

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

public sealed class PdfElementCompose {
    private readonly PdfDoc _doc;
    internal PdfElementCompose(PdfDoc doc) { _doc = doc; }
    public PdfElementCompose H1(string text) { _doc.H1(text); return this; }
    public PdfElementCompose H2(string text) { _doc.H2(text); return this; }
    public PdfElementCompose H3(string text) { _doc.H3(text); return this; }
    public PdfElementCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
}

