namespace OfficeIMO.Pdf;

internal sealed class NumberedListBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<string> Items { get; }
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
        var snapshot = new System.Collections.Generic.List<string>();
        snapshot.AddRange(items.Where(item => item is not null));
        Items = snapshot.AsReadOnly();
    }
}
