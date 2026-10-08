namespace OfficeIMO.Pdf;

internal static class PdfPageNumberFormatter {
    internal static string Format(int number, PdfPageNumberStyle style) {
        Guard.PageNumberStyle(style, nameof(style));
        if (number < 1)
            throw new ArgumentOutOfRangeException(nameof(number), "PDF page number must be positive.");
        OfficeIMO.Core.OfficeNumberStyle commonStyle = style switch {
            PdfPageNumberStyle.Arabic => OfficeIMO.Core.OfficeNumberStyle.Decimal,
            PdfPageNumberStyle.LowerRoman => OfficeIMO.Core.OfficeNumberStyle.LowerRoman,
            PdfPageNumberStyle.UpperRoman => OfficeIMO.Core.OfficeNumberStyle.UpperRoman,
            PdfPageNumberStyle.LowerLetter => OfficeIMO.Core.OfficeNumberStyle.LowerLetter,
            PdfPageNumberStyle.UpperLetter => OfficeIMO.Core.OfficeNumberStyle.UpperLetter,
            _ => throw new ArgumentException("Unsupported PDF page number style.", nameof(style))
        };
        return OfficeIMO.Core.OfficeNumberFormatter.Format(number, commonStyle);
    }
}
