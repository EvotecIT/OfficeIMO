namespace OfficeIMO.Pdf;

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public PdfAlign Align { get; }
    public ParagraphBlock(string text, PdfAlign align) { Text = text; Align = align; }
}
