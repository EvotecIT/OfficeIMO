namespace OfficeIMO.Pdf;

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

