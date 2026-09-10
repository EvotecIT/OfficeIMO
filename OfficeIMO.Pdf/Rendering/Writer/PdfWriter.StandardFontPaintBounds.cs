using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Vertical FontBBox values from Adobe's Core 14 AFMs, in 1/1000 em units:
    // https://download.macromedia.com/pub/developer/opentype/tech-notes/Core14_AFMs.zip
    // Unembedded fonts have no outlines to measure. The whole-font box is a
    // conservative fallback; line ascenders/descenders omit some accented glyphs.
    private static OfficeTextPaintBounds GetStandardFontPaintBounds(PdfStandardFont font, double size) {
        (int Bottom, int Top) bounds = font switch {
            PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique => (-225, 931),
            PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaBoldOblique => (-228, 962),
            PdfStandardFont.TimesRoman => (-218, 898),
            PdfStandardFont.TimesItalic => (-217, 883),
            PdfStandardFont.TimesBold => (-218, 935),
            PdfStandardFont.TimesBoldItalic => (-218, 921),
            PdfStandardFont.Courier or PdfStandardFont.CourierOblique => (-250, 805),
            PdfStandardFont.CourierBold or PdfStandardFont.CourierBoldOblique => (-250, 801),
            _ => throw new System.ArgumentOutOfRangeException(nameof(font))
        };
        return new OfficeTextPaintBounds(-bounds.Top * size / 1000D, -bounds.Bottom * size / 1000D);
    }
}
