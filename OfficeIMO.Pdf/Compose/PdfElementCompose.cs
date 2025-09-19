namespace OfficeIMO.Pdf;

/// <summary>Builder for nested elements used within item builders.</summary>
public class PdfElementCompose {
    private readonly PdfDoc _doc;
    internal PdfElementCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Adds an H1 heading.</summary>
    /// <param name="text">Heading text.</param>
    public PdfElementCompose H1(string text) { _doc.H1(text); return this; }
    /// <summary>Adds an H2 heading.</summary>
    /// <param name="text">Heading text.</param>
    public PdfElementCompose H2(string text) { _doc.H2(text); return this; }
    /// <summary>Adds an H3 heading.</summary>
    /// <param name="text">Heading text.</param>
    public PdfElementCompose H3(string text) { _doc.H3(text); return this; }
    /// <summary>Adds a paragraph built from styled text runs.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfElementCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    /// <summary>Adds a simple text table.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
}
