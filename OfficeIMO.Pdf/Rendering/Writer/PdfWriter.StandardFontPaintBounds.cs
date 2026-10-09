using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // FontBBox values from Adobe's Core 14 AFMs, in 1/1000 em units:
    // https://download.macromedia.com/pub/developer/opentype/tech-notes/Core14_AFMs.zip
    // Use the painted WinAnsi glyphs rather than the font's largest accent or
    // descender. Unavailable glyph metadata retains the conservative whole-font box.
    private static OfficeTextPaintBounds GetStandardFontPaintBounds(string? text, PdfStandardFont font, double size) {
        int top = int.MinValue, bottom = int.MaxValue;
        foreach (char value in text ?? string.Empty) {
            if (value is ' ' or '\u00A0') continue;
            int glyphTop, glyphBottom;
            if (!PdfWinAnsiEncoding.CanEncodeCharacter(value) ||
                !PdfStandardFontWidths.TryGetVerticalBounds1000(font, value, out glyphBottom, out glyphTop)) {
                var bounds = GetStandardFontBoundingBox(font);
                glyphTop = bounds.Top; glyphBottom = bounds.Bottom;
            }
            top = Math.Max(top, glyphTop); bottom = Math.Min(bottom, glyphBottom);
        }
        return top == int.MinValue ? default :
            new OfficeTextPaintBounds(-top * size / 1000D, -bottom * size / 1000D);
    }

    private static (double Left, double Right) GetStandardFontHorizontalPaintBounds(
        PdfStandardFont font, double size, double advance) {
        var bounds = GetStandardFontBoundingBox(font);
        // Glyph origins lie within the run advance. Enclose their whole-font boxes
        // when the unembedded font has no individual outlines to measure.
        return (Math.Min(0D, bounds.Left * size / 1000D),
            advance + Math.Max(0D, bounds.Right * size / 1000D));
    }

    private static (int Left, int Bottom, int Right, int Top) GetStandardFontBoundingBox(PdfStandardFont font) => font switch {
        PdfStandardFont.Helvetica => (-166, -225, 1000, 931),
        PdfStandardFont.HelveticaOblique => (-170, -225, 1116, 931),
        PdfStandardFont.HelveticaBold => (-170, -228, 1003, 962),
        PdfStandardFont.HelveticaBoldOblique => (-174, -228, 1114, 962),
        PdfStandardFont.TimesRoman => (-168, -218, 1000, 898),
        PdfStandardFont.TimesItalic => (-169, -217, 1010, 883),
        PdfStandardFont.TimesBold => (-168, -218, 1000, 935),
        PdfStandardFont.TimesBoldItalic => (-200, -218, 996, 921),
        PdfStandardFont.Courier => (-23, -250, 715, 805),
        PdfStandardFont.CourierOblique => (-27, -250, 849, 805),
        PdfStandardFont.CourierBold => (-113, -250, 749, 801),
        PdfStandardFont.CourierBoldOblique => (-57, -250, 869, 801),
        _ => throw new System.ArgumentOutOfRangeException(nameof(font))
    };
}
