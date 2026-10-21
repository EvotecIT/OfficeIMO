using System;
using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static ContentStreamBuilder WriteLogicalTextAnchor(StringBuilder sb, PdfOptions options, PdfStandardFont font, string fontResource, double anchorX, double anchorY,
        string actualText, double width, double height) {
        bool hasBounds = width > 0D && height > 0D;
        double fontSize = hasBounds ? height : 1D;
        int anchorCount = hasBounds ? CountLogicalAnchorScalars(actualText) : 1;
        PdfTextShowCommand anchor = hasBounds
            ? EncodeBoundedLogicalTextAnchor(font, options, anchorCount)
            : EncodeActualTextAnchor(font, options, anchorCount);
        var content = new ContentStreamBuilder(sb)
            .SaveState()
            .BeginText()
            // Bounds-derived scaling uses this same fractional height. Keep
            // its precision so tiny carriers do not acquire a longer baseline.
            .Font(fontResource, fontSize, preserveLogicalPrecision: hasBounds)
            .TextRenderingMode(3)
            .TextMatrix(anchorX, anchorY);
        if (hasBounds) {
            // The configured embedded font can have different space metrics
            // from its standard-font fallback. Use the encoded run's advance.
            content.WordSpacing(0D).TextRise(0D).HorizontalTextScaling(ResolveLogicalAnchorScaling(anchor, width, fontSize));
        }
        content.ShowText(anchor, fontSize);
        return content.EndText();
    }

    private static PdfTextShowCommand EncodeBoundedLogicalTextAnchor(PdfStandardFont font, PdfOptions options, int count) {
        PdfTextShowCommand anchor = EncodeActualTextAnchor(font, options, count);
        // Font positioning styles painted glyphs, not the caller's explicit
        // semantic rectangle. Emit one nominal run so its first text-show
        // operation already carries the complete replacement geometry.
        double? nominalAdvance = anchor.PositionedGlyphs == null
            ? anchor.AdvanceWidth1000
            : anchor.PositionedGlyphs.Sum(glyph => (double)glyph.NominalWidth1000);
        return new PdfTextShowCommand(anchor.GlyphHex,
            advanceWidth1000: nominalAdvance, wordSpaceCount: anchor.WordSpaceCount, glyphCount: anchor.GlyphCount,
            unitsPerEm: anchor.UnitsPerEm, fontMetricScale: anchor.FontMetricScale);
    }

    private static double ResolveLogicalAnchorScaling(PdfTextShowCommand anchor, double width, double height) {
        double advance = anchor.AdvanceWidth1000.GetValueOrDefault() * height / 1000D;
        if (advance <= 0D) throw new InvalidOperationException("The logical text anchor font requires a positive space advance.");
        return width / advance * 100D;
    }

    private static int CountLogicalAnchorScalars(string text) {
        int count = 0;
        for (int index = 0; index < text.Length; index++, count++) {
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length &&
                char.IsLowSurrogate(text[index + 1])) index++;
        }
        return count;
    }
}
