namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static PdfStandardFont ChooseNormal(PdfStandardFont requested) => requested switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => PdfStandardFont.Helvetica,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => PdfStandardFont.TimesRoman,
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => PdfStandardFont.Courier,
        _ => PdfStandardFont.Courier
    };

    private static PdfStandardFont ChooseBold(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaBold,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold => PdfStandardFont.TimesBold,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold => PdfStandardFont.CourierBold,
        _ => PdfStandardFont.CourierBold
    };

    private static PdfStandardFont ChooseItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => PdfStandardFont.TimesItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => PdfStandardFont.CourierOblique,
        _ => PdfStandardFont.HelveticaOblique
    };

    private static PdfStandardFont ChooseBoldItalic(PdfStandardFont normal) => normal switch {
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBold => PdfStandardFont.HelveticaBoldOblique,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBold => PdfStandardFont.TimesBoldItalic,
        PdfStandardFont.Courier or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBold => PdfStandardFont.CourierBoldOblique,
        _ => PdfStandardFont.HelveticaBoldOblique
    };

    private static double GlyphWidthEmFor(PdfStandardFont font) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => 0.6,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold or PdfStandardFont.HelveticaOblique or PdfStandardFont.HelveticaBoldOblique => 0.55,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => 0.5,
        _ => 0.6
    };

    private static double GetDescender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => fontSize * 0.23,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => fontSize * 0.22,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => fontSize * 0.26,
        _ => fontSize * 0.23
    };

    private static double GetAscender(PdfStandardFont font, double fontSize) => font switch {
        PdfStandardFont.Courier or PdfStandardFont.CourierBold => fontSize * 0.72,
        PdfStandardFont.Helvetica or PdfStandardFont.HelveticaBold => fontSize * 0.74,
        PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold => fontSize * 0.72,
        _ => fontSize * 0.72
    };
}

