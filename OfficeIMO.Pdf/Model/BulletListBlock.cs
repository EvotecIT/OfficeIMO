namespace OfficeIMO.Pdf;

internal sealed class BulletListBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<string> Items { get; }
    public System.Collections.Generic.IReadOnlyList<PdfListItem> RichItems { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public PdfListStyle? Style { get; }
    public BulletListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        Guard.LeftCenterRightAlign(align, nameof(align), "Bullet list");
        Align = align;
        Color = color;
        Style = style?.Clone();
        var richSnapshot = new System.Collections.Generic.List<PdfListItem>();
        foreach (string? item in items) {
            if (item != null) {
                richSnapshot.Add(new PdfListItem(item));
            }
        }

        RichItems = richSnapshot.AsReadOnly();
        Items = richSnapshot.Select(item => item.Text).ToList().AsReadOnly();
    }

    public BulletListBlock(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align, PdfColor? color, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        Guard.LeftCenterRightAlign(align, nameof(align), "Bullet list");
        Align = align;
        Color = color;
        Style = style?.Clone();
        var richSnapshot = new System.Collections.Generic.List<PdfListItem>();
        foreach (PdfListItem? item in items) {
            if (item != null) {
                richSnapshot.Add(new PdfListItem(item.Runs, item.BookmarkName, item.Marker));
            }
        }

        RichItems = richSnapshot.AsReadOnly();
        Items = richSnapshot.Select(item => item.Text).ToList().AsReadOnly();
    }
}
