namespace OfficeIMO.Pdf;

/// <summary>
/// Canonical internal model shared by bulleted and numbered lists.
/// </summary>
internal abstract class PdfListBlock : IPdfBlock {
    protected PdfListBlock(
        System.Collections.Generic.IEnumerable<string> items,
        PdfAlign align,
        PdfColor? color,
        PdfListStyle? style,
        string validationName) {
        Guard.NotNull(items, nameof(items));
        Guard.LeftCenterRightAlign(align, nameof(align), validationName);

        var richItems = new System.Collections.Generic.List<PdfListItem>();
        foreach (string? item in items) {
            if (item != null) {
                richItems.Add(new PdfListItem(item));
            }
        }

        Align = align;
        Color = color;
        Style = style?.Clone();
        RichItems = richItems.AsReadOnly();
        Items = CreateTextSnapshot(richItems);
    }

    protected PdfListBlock(
        System.Collections.Generic.IEnumerable<PdfListItem> items,
        PdfAlign align,
        PdfColor? color,
        PdfListStyle? style,
        string validationName) {
        Guard.NotNull(items, nameof(items));
        Guard.LeftCenterRightAlign(align, nameof(align), validationName);

        var richItems = new System.Collections.Generic.List<PdfListItem>();
        foreach (PdfListItem? item in items) {
            if (item != null) {
                richItems.Add(new PdfListItem(item.Runs, item.BookmarkName, item.Marker));
            }
        }

        Align = align;
        Color = color;
        Style = style?.Clone();
        RichItems = richItems.AsReadOnly();
        Items = CreateTextSnapshot(richItems);
    }

    public System.Collections.Generic.IReadOnlyList<string> Items { get; }

    public System.Collections.Generic.IReadOnlyList<PdfListItem> RichItems { get; }

    public PdfAlign Align { get; }

    public PdfColor? Color { get; }

    public PdfListStyle? Style { get; }

    internal abstract bool IsNumbered { get; }

    internal abstract int StartingNumber { get; }

    internal abstract double PreferredMarkerWidthFactor { get; }

    internal abstract string GetDefaultMarker(int itemIndex);

    internal string GetMarker(int itemIndex) =>
        RichItems[itemIndex].Marker ?? GetDefaultMarker(itemIndex);

    internal string GetWidestDefaultMarker() =>
        GetDefaultMarker(System.Math.Max(0, RichItems.Count - 1));

    internal PdfAlign GetMarkerAlign(PdfListStyle? style) =>
        style?.MarkerAlign ?? (IsNumbered ? PdfAlign.Right : PdfAlign.Left);

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> CreateTextSnapshot(
        System.Collections.Generic.List<PdfListItem> items) {
        var text = new System.Collections.Generic.List<string>(items.Count);
        for (int index = 0; index < items.Count; index++) {
            text.Add(items[index].Text);
        }

        return text.AsReadOnly();
    }
}
