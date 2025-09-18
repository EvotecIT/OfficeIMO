namespace OfficeIMO.Pdf;

internal sealed class BulletListBlock : IPdfBlock {
    public System.Collections.Generic.List<string> Items { get; } = new();
    public PdfAlign Align { get; }
    public BulletListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align) { Align = align; Items.AddRange(items); }
}

