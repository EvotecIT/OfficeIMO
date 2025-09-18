namespace OfficeIMO.Pdf;

internal interface IPdfBlock { }

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public HeadingBlock(int level, string text) { Level = level; Text = text; }
}

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public ParagraphBlock(string text) { Text = text; }
}

internal sealed class PageBreakBlock : IPdfBlock { }

