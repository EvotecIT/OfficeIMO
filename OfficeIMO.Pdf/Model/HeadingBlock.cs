namespace OfficeIMO.Pdf;

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public string? LinkUri { get; }
    public HeadingBlock(int level, string text, PdfAlign align, PdfColor? color, string? linkUri = null) { Level = level; Text = text; Align = align; Color = color; LinkUri = linkUri; }
}
