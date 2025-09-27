namespace OfficeIMO.Pdf;

/// <summary>Builder for individual flow items (headings, paragraphs, tables, images).</summary>
public class PdfItemCompose {
    private readonly PdfDoc _doc;
    internal PdfItemCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Starts a new page.</summary>
    public PdfItemCompose PageBreak() { _doc.PageBreak(); return this; }
    /// <summary>Adds an H1 heading.</summary>
    public PdfItemCompose H1(string text) { _doc.H1(text); return this; }
    /// <summary>Adds an H2 heading.</summary>
    public PdfItemCompose H2(string text) { _doc.H2(text); return this; }
    /// <summary>Adds an H3 heading.</summary>
    public PdfItemCompose H3(string text) { _doc.H3(text); return this; }
    /// <summary>Adds a paragraph built from styled text runs.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfItemCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    /// <summary>Adds a simple text table.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfItemCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    /// <summary>Builds nested elements (e.g., grouping heading + paragraph).</summary>
    /// <param name="build">Delegate composing nested elements.</param>
    public PdfItemCompose Element(System.Action<PdfElementCompose> build) { var el = new PdfElementCompose(_doc); build(el); return this; }
    /// <summary>Adds a horizontal rule.</summary>
    /// <param name="thickness">Line thickness (pt).</param>
    /// <param name="color">Optional color; gray by default.</param>
    /// <param name="spacingBefore">Top spacing (pt).</param>
    /// <param name="spacingAfter">Bottom spacing (pt).</param>
    public PdfItemCompose HR(double thickness = 0.5, PdfColor? color = null, double spacingBefore = 6, double spacingAfter = 6) { _doc.HR(thickness, color, spacingBefore, spacingAfter); return this; }
    /// <summary>Adds a paragraph inside a styled panel.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfItemCompose PanelParagraph(System.Action<PdfParagraphBuilder> build, PanelStyle style, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.PanelParagraph(build, style, align, defaultColor); return this; }
    /// <summary>Adds an image from JPEG bytes.</summary>
    /// <param name="jpegBytes">JPEG-encoded image bytes.</param>
    /// <param name="width">Target width in points.</param>
    /// <param name="height">Target height in points.</param>
    /// <param name="align">Image alignment inside content width.</param>
    public PdfItemCompose Image(byte[] jpegBytes, double width, double height, PdfAlign align = PdfAlign.Left) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        _doc.Image(jpegBytes, width, height, align);
        return this;
    }
}
