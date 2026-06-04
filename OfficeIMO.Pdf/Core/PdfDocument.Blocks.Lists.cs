using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds a simple bullet list.</summary>
    public PdfDocument Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }

    /// <summary>Adds a bullet list whose items can contain rich inline text runs.</summary>
    public PdfDocument RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }

    /// <summary>Adds a simple numbered list.</summary>
    public PdfDocument Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }

    /// <summary>Adds a numbered list whose items can contain rich inline text runs.</summary>
    public PdfDocument RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }

    /// <summary>Sets the document-wide default style for bullet and numbered lists.</summary>
    public PdfDocument DefaultListStyle(PdfListStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultListStyle = style;
        return this;
    }
}
