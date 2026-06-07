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

    internal static string BuildGeneratedPageLabelsDictionary(IReadOnlyList<PdfPageLabelRange> ranges) {
        Guard.NotNull(ranges, nameof(ranges));
        if (ranges.Count == 0) {
            throw new ArgumentException("At least one PDF page-label range is required.", nameof(ranges));
        }

        var ordered = ranges.OrderBy(range => range.StartPageNumber).ToList();
        var seenStartPages = new HashSet<int>();
        var sb = new StringBuilder();
        sb.Append("<< /Nums [");

        for (int i = 0; i < ordered.Count; i++) {
            PdfPageLabelRange range = ordered[i];
            if (!seenStartPages.Add(range.StartPageNumber)) {
                throw new ArgumentException("PDF page-label ranges cannot contain duplicate start pages.", nameof(ranges));
            }

            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append((range.StartPageNumber - 1).ToString(CultureInfo.InvariantCulture))
                .Append(" << /S /")
                .Append(GetStyleName(range.Style))
                .Append(" /St ")
                .Append(range.StartNumber.ToString(CultureInfo.InvariantCulture));

            if (range.Prefix != null) {
                sb.Append(" /P ")
                    .Append(PdfSyntaxEscaper.TextString(range.Prefix));
            }

            sb.Append(" >>");
        }

        sb.Append("] >>\n");
        return sb.ToString();
    }
}
