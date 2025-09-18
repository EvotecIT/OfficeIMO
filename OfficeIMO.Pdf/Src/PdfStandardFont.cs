namespace OfficeIMO.Pdf;

/// <summary>
/// Standard 14 PDF fonts supported by all PDF viewers without embedding.
/// </summary>
public enum PdfStandardFont {
    Helvetica,
    HelveticaBold,
    TimesRoman,
    TimesBold,
    Courier,
    CourierBold
}

internal static class PdfFontNames {
    internal static string ToBaseFontName(this PdfStandardFont f) => f switch {
        PdfStandardFont.Helvetica => "Helvetica",
        PdfStandardFont.HelveticaBold => "Helvetica-Bold",
        PdfStandardFont.TimesRoman => "Times-Roman",
        PdfStandardFont.TimesBold => "Times-Bold",
        PdfStandardFont.Courier => "Courier",
        PdfStandardFont.CourierBold => "Courier-Bold",
        _ => "Helvetica"
    };
}

