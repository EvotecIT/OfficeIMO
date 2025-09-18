namespace OfficeIMO.Pdf;

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public ParagraphBlock(string text, PdfAlign align, PdfColor? color) { Text = text; Align = align; Color = color; }
}
