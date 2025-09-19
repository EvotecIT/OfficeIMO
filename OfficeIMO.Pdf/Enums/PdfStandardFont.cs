namespace OfficeIMO.Pdf;

/// <summary>
/// Standard 14 PDF Type 1 fonts supported by all PDF viewers without embedding.
/// </summary>
public enum PdfStandardFont {
    /// <summary>Helvetica (regular).</summary>
    Helvetica,
    /// <summary>Helvetica-Oblique (italic).</summary>
    HelveticaOblique,
    /// <summary>Helvetica-Bold.</summary>
    HelveticaBold,
    /// <summary>Helvetica-BoldOblique (bold italic).</summary>
    HelveticaBoldOblique,
    /// <summary>Times-Roman (regular).</summary>
    TimesRoman,
    /// <summary>Times-Italic.</summary>
    TimesItalic,
    /// <summary>Times-Bold.</summary>
    TimesBold,
    /// <summary>Times-BoldItalic.</summary>
    TimesBoldItalic,
    /// <summary>Courier (regular, monospaced).</summary>
    Courier,
    /// <summary>Courier-Oblique.</summary>
    CourierOblique,
    /// <summary>Courier-Bold (monospaced).</summary>
    CourierBold,
    /// <summary>Courier-BoldOblique (monospaced).</summary>
    CourierBoldOblique
}

internal static class PdfFontNames {
    internal static string ToBaseFontName(this PdfStandardFont f) => f switch {
        PdfStandardFont.Helvetica => "Helvetica",
        PdfStandardFont.HelveticaOblique => "Helvetica-Oblique",
        PdfStandardFont.HelveticaBold => "Helvetica-Bold",
        PdfStandardFont.HelveticaBoldOblique => "Helvetica-BoldOblique",
        PdfStandardFont.TimesRoman => "Times-Roman",
        PdfStandardFont.TimesItalic => "Times-Italic",
        PdfStandardFont.TimesBold => "Times-Bold",
        PdfStandardFont.TimesBoldItalic => "Times-BoldItalic",
        PdfStandardFont.Courier => "Courier",
        PdfStandardFont.CourierOblique => "Courier-Oblique",
        PdfStandardFont.CourierBold => "Courier-Bold",
        PdfStandardFont.CourierBoldOblique => "Courier-BoldOblique",
        _ => "Helvetica"
    };
}
