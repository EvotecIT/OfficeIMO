namespace OfficeIMO.Pdf;

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public HeadingBlock(int level, string text) { Level = level; Text = text; }
}

