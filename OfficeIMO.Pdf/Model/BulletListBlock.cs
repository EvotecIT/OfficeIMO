namespace OfficeIMO.Pdf;

internal sealed class BulletListBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<string> Items { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public PdfListStyle? Style { get; }
    public BulletListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        Guard.LeftCenterRightAlign(align, nameof(align), "Bullet list");
        Align = align;
        Color = color;
        Style = style?.Clone();
        var snapshot = new System.Collections.Generic.List<string>();
        snapshot.AddRange(items.Where(item => item is not null));
        Items = snapshot.AsReadOnly();
    }
}
