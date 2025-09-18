namespace OfficeIMO.Pdf;

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public PdfAlign Align { get; }
    public HeadingBlock(int level, string text, PdfAlign align) { Level = level; Text = text; Align = align; }
}
