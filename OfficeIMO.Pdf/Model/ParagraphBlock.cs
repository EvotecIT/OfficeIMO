namespace OfficeIMO.Pdf;

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public PdfAlign Align { get; }
    public global::OfficeIMO.OfficeColor? Color { get; }
    public ParagraphBlock(string text, PdfAlign align, global::OfficeIMO.OfficeColor? color) { Text = text; Align = align; Color = color; }
}
