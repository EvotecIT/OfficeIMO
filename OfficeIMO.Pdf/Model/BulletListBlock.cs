namespace OfficeIMO.Pdf;

internal sealed class BulletListBlock : PdfListBlock {
    public BulletListBlock(System.Collections.Generic.IEnumerable<string> items, PdfAlign align, PdfColor? color, PdfListStyle? style = null)
        : base(items, align, color, style, "Bullet list") { }

    public BulletListBlock(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align, PdfColor? color, PdfListStyle? style = null)
        : base(items, align, color, style, "Bullet list") { }

    internal override bool IsNumbered => false;

    internal override int StartingNumber => 1;

    internal override double PreferredMarkerWidthFactor => 1.5D;

    internal override string GetDefaultMarker(int itemIndex) => "•";
}
