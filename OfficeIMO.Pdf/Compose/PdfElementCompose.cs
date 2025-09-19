namespace OfficeIMO.Pdf;

/// <summary>Builder for nested elements used within item builders.</summary>
public class PdfElementCompose {
    private readonly PdfDoc _doc;
    internal PdfElementCompose(PdfDoc doc) { _doc = doc; }
    public PdfElementCompose H1(string text) { _doc.H1(text); return this; }
    public PdfElementCompose H2(string text) { _doc.H2(text); return this; }
    public PdfElementCompose H3(string text) { _doc.H3(text); return this; }
    public PdfElementCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Paragraph(build, align, defaultColor); return this; }
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
}

