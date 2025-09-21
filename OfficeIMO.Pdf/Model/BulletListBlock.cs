namespace OfficeIMO.Pdf;

internal sealed class BulletListBlock : IPdfBlock {
    public System.Collections.Generic.List<string> Items { get; } = new();
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public BulletListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color) {
        Guard.NotNull(items, nameof(items));
        Align = align;
        Color = color;
        Items.AddRange(items.Where(item => item is not null));
    }
}
