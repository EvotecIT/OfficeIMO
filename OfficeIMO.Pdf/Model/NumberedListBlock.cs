namespace OfficeIMO.Pdf;

internal sealed class NumberedListBlock : PdfListBlock {
    public int StartNumber { get; }

    public NumberedListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color, int startNumber, PdfListStyle? style = null)
        : base(Validate(items, startNumber), align, color, style, "Numbered list") {
        StartNumber = startNumber;
    }

    public NumberedListBlock(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align, PdfColor? color, int startNumber, PdfListStyle? style = null)
        : base(Validate(items, startNumber), align, color, style, "Numbered list") {
        StartNumber = startNumber;
    }

    internal override bool IsNumbered => true;

    internal override int StartingNumber => StartNumber;

    internal override double PreferredMarkerWidthFactor => 2D;

    internal override string GetDefaultMarker(int itemIndex) =>
        (StartNumber + itemIndex).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".";

    private static System.Collections.Generic.IEnumerable<T> Validate<T>(
        System.Collections.Generic.IEnumerable<T> items,
        int startNumber) {
        Guard.NotNull(items, nameof(items));
        if (startNumber < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(startNumber), "Numbered lists must start at 1 or greater.");
        }

        return items;
    }
}
