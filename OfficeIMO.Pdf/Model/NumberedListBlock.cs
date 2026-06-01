namespace OfficeIMO.Pdf;

internal sealed class NumberedListBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<string> Items { get; }
    public System.Collections.Generic.IReadOnlyList<PdfListItem> RichItems { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public int StartNumber { get; }
    public PdfListStyle? Style { get; }

    public NumberedListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color, int startNumber, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        if (startNumber < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(startNumber), "Numbered lists must start at 1 or greater.");
        }

        Guard.LeftCenterRightAlign(align, nameof(align), "Numbered list");
        Align = align;
        Color = color;
        StartNumber = startNumber;
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

    public NumberedListBlock(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align, PdfColor? color, int startNumber, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        if (startNumber < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(startNumber), "Numbered lists must start at 1 or greater.");
        }

        Guard.LeftCenterRightAlign(align, nameof(align), "Numbered list");
        Align = align;
        Color = color;
        StartNumber = startNumber;
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
