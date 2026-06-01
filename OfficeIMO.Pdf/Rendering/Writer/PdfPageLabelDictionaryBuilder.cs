using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfPageLabelDictionaryBuilder {
    internal static string BuildGeneratedPageLabelsDictionary(PdfPageNumberStyle style, int startNumber, string? prefix = null) {
        Guard.PageNumberStyle(style, nameof(style));
        if (startNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(startNumber), "PDF page-label start number must be positive.");
        }

        if (prefix != null) {
            ValidatePrefix(prefix, nameof(prefix));
        }

        var sb = new StringBuilder();
        sb.Append("<< /Nums [0 << /S /")
            .Append(GetStyleName(style))
            .Append(" /St ")
            .Append(startNumber.ToString(CultureInfo.InvariantCulture));

        if (prefix != null) {
            sb.Append(" /P ")
                .Append(PdfSyntaxEscaper.TextString(prefix));
        }

        sb.Append(" >>] >>\n");
        return sb.ToString();
    }

    internal static void ValidatePrefix(string? prefix, string paramName) {
        if (prefix == null) {
            return;
        }

        if (string.IsNullOrWhiteSpace(prefix)) {
            throw new ArgumentException("PDF page-label prefix cannot be empty or whitespace.", paramName);
        }

        for (int i = 0; i < prefix.Length; i++) {
            if (char.IsControl(prefix[i])) {
                throw new ArgumentException("PDF page-label prefix cannot contain control characters.", paramName);
            }
        }
    }

    internal static string GetStyleName(PdfPageNumberStyle style) {
        Guard.PageNumberStyle(style, nameof(style));
        return style switch {
            PdfPageNumberStyle.Arabic => "D",
            PdfPageNumberStyle.LowerRoman => "r",
            PdfPageNumberStyle.UpperRoman => "R",
            PdfPageNumberStyle.LowerLetter => "a",
            PdfPageNumberStyle.UpperLetter => "A",
            _ => throw new ArgumentOutOfRangeException(nameof(style), "PDF page-label style is not supported.")
        };
    }
}
