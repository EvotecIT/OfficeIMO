namespace OfficeIMO.Pdf;

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public OfficeAlignment Align { get; }
    public global::OfficeIMO.OfficeColor? Color { get; }
    public ParagraphBlock(string text, OfficeAlignment align, global::OfficeIMO.OfficeColor? color) { Text = text; Align = align; Color = color; }
}
