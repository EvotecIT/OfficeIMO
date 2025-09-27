namespace OfficeIMO.Pdf;

/// <summary>Column content builder used within <see cref="PdfRowCompose"/>.</summary>
public class PdfRowColumnCompose {
    private readonly RowColumn _col;
    internal PdfRowColumnCompose(RowColumn col) { _col = col; }
    /// <summary>Adds an H1 heading in the column.</summary>
    public PdfRowColumnCompose H1(string text) { _col.Blocks.Add(new HeadingBlock(1, text, PdfAlign.Left, null)); return this; }
    /// <summary>Adds an H2 heading in the column.</summary>
    public PdfRowColumnCompose H2(string text) { _col.Blocks.Add(new HeadingBlock(2, text, PdfAlign.Left, null)); return this; }
    /// <summary>Adds an H3 heading in the column.</summary>
    public PdfRowColumnCompose H3(string text) { _col.Blocks.Add(new HeadingBlock(3, text, PdfAlign.Left, null)); return this; }
    /// <summary>Adds a paragraph built from styled runs to the column.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfRowColumnCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        var b = new PdfParagraphBuilder(align, defaultColor);
        build(b);
        _col.Blocks.Add(new RichParagraphBlock(b.Build().Runs, align, defaultColor));
        return this;
    }
    /// <summary>Adds a horizontal rule in the column.</summary>
    public PdfRowColumnCompose HR(double thickness = 0.5, PdfColor? color = null, double spacingBefore = 6, double spacingAfter = 6) { _col.Blocks.Add(new HorizontalRuleBlock(thickness, color ?? PdfColor.Gray, spacingBefore, spacingAfter)); return this; }
    /// <summary>Adds a JPEG image in the column.</summary>
    public PdfRowColumnCompose Image(byte[] jpegBytes, double width, double height, PdfAlign align = PdfAlign.Left) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        PdfDoc.WarnIfBytesNotJpeg(jpegBytes);

        _col.Blocks.Add(new ImageBlock(jpegBytes, width, height, align));
        return this;
    }
}
