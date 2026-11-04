#pragma warning disable RS0026 // Plain and rich list items share the list operation names and optional layout parameters.
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Appends a bulleted list of plain items.</summary>
    public PdfList AddList(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        return AddListBlock(new BulletListBlock(items, align, color, style));
    }

    /// <summary>Appends a bulleted list whose items can contain rich inline text runs.</summary>
    public PdfList AddList(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        return AddListBlock(new BulletListBlock(items, align, color, style));
    }

    /// <summary>Appends a numbered list of plain items.</summary>
    public PdfList AddNumberedList(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        return AddListBlock(new NumberedListBlock(items, align, color, startNumber, style));
    }

    /// <summary>Appends a numbered list whose items can contain rich inline text runs.</summary>
    public PdfList AddNumberedList(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        return AddListBlock(new NumberedListBlock(items, align, color, startNumber, style));
    }

    private PdfList AddListBlock(PdfListBlock list) {
        AddBlock(list);
        return new PdfList(list.RichItems.Count, list.IsNumbered);
    }

    // Chaining forms used by adapters and tests that compose a document in one expression.
    internal PdfDocument Bullets(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) { AddList(items, align, color, style); return this; }
    internal PdfDocument RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) { AddList(items, align, color, style); return this; }
    internal PdfDocument Numbered(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) { AddNumberedList(items, align, color, startNumber, style); return this; }
    internal PdfDocument RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) { AddNumberedList(items, align, color, startNumber, style); return this; }

    /// <summary>Sets the document-wide default style for bullet and numbered lists.</summary>
    internal PdfDocument DefaultListStyle(PdfListStyle style) {
        Guard.NotNull(style, nameof(style));
        OptionsToChange.DefaultListStyle = style;
        return this;
    }
}
