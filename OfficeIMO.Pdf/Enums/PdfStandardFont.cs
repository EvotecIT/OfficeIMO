namespace OfficeIMO.Pdf;

/// <summary>
/// Standard 14 PDF Type 1 fonts supported by all PDF viewers without embedding.
/// </summary>
public enum PdfStandardFont {
    /// <summary>Helvetica (regular).</summary>
    Helvetica,
    /// <summary>Helvetica-Bold.</summary>
    HelveticaBold,
    /// <summary>Times-Roman (regular).</summary>
    TimesRoman,
    /// <summary>Times-Bold.</summary>
    TimesBold,
    /// <summary>Courier (regular, monospaced).</summary>
    Courier,
    /// <summary>Courier-Bold (monospaced).</summary>
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
