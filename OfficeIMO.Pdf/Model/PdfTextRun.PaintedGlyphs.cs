namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Text contains already-shaped glyph characters in left-to-right paint order.</summary>
    internal bool PreservesPaintedGlyphs { get; private set; }

    /// <summary>
    /// Preserves an upstream layout owner's contextual glyph forms and visual order.
    /// Logical source text remains the upstream owner's separate semantic contract.
    /// </summary>
    internal PdfTextRun WithPaintedGlyphPreservation(bool preservePaintedGlyphs = true) {
        if (InlineElement != null || PreservesPaintedGlyphs == preservePaintedGlyphs) return this;
        var copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic, Strike,
            FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader,
            TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.PreservesPaintedGlyphs = preservePaintedGlyphs;
        return copy;
    }
}
