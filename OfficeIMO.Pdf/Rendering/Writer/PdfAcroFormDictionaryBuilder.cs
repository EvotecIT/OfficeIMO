namespace OfficeIMO.Pdf;

internal static class PdfAcroFormDictionaryBuilder {
    internal static string BuildAcroFormDictionary(IReadOnlyList<int> fieldObjectIds, int helveticaFontId) {
        Guard.NotNull(fieldObjectIds, nameof(fieldObjectIds));
        if (fieldObjectIds.Count == 0) {
            throw new ArgumentException("PDF AcroForm dictionary requires at least one field object.", nameof(fieldObjectIds));
        }

        var sb = new StringBuilder();
        sb.Append("<< /Fields [");
        for (int i = 0; i < fieldObjectIds.Count; i++) {
            sb.Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(fieldObjectIds[i]));
        }

        sb.Append(" ] /NeedAppearances true /DR << /Font << /Helv ")
            .Append(PdfSyntaxEscaper.IndirectReference(helveticaFontId))
            .Append(" >> >> /DA (/Helv 10 Tf 0 g) >>\n");
        return sb.ToString();
    }

    internal static string BuildTextFieldAppearanceStreamDictionary(double width, double height, int helveticaFontId, int contentLength) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        if (contentLength < 0) {
            throw new ArgumentOutOfRangeException(nameof(contentLength), "PDF appearance stream length cannot be negative.");
        }

        return "<< /Type /XObject /Subtype /Form /BBox [0 0 " +
            Format(width) +
            " " +
            Format(height) +
            "] /Resources << /Font << /Helv " +
            PdfSyntaxEscaper.IndirectReference(helveticaFontId) +
            " >> >> /Length " +
            contentLength.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>";
    }

    internal static string BuildCheckBoxAppearanceStreamDictionary(double width, double height, int contentLength) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        if (contentLength < 0) {
            throw new ArgumentOutOfRangeException(nameof(contentLength), "PDF appearance stream length cannot be negative.");
        }

        return "<< /Type /XObject /Subtype /Form /BBox [0 0 " +
            Format(width) +
            " " +
            Format(height) +
            "] /Length " +
            contentLength.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>";
    }

    internal static string BuildTextFieldAppearanceContent(double width, double height, string value, double fontSize) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(value, nameof(value));
        Guard.Positive(fontSize, nameof(fontSize));

        double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
        double textX = 3D;
        double textWidth = Math.Max(0D, width - 6D);
        string clippedValue = value;
        if (textWidth <= 0.001D) {
            clippedValue = string.Empty;
        }

        return "q\n" +
            "1 1 1 rg 0 0 " + Format(width) + " " + Format(height) + " re f\n" +
            "0.75 G 1 w 0.5 0.5 " + Format(Math.Max(0D, width - 1D)) + " " + Format(Math.Max(0D, height - 1D)) + " re S\n" +
            "BT /Helv " + Format(fontSize) + " Tf 0 g " + Format(textX) + " " + Format(baseline) + " Td " + PdfSyntaxEscaper.WinAnsiHexString(clippedValue) + " Tj ET\n" +
            "Q\n";
    }

    internal static string BuildCheckBoxAppearanceContent(double width, double height, bool selected) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        double boxWidth = Math.Max(0D, width - 1D);
        double boxHeight = Math.Max(0D, height - 1D);
        string content =
            "q\n" +
            "1 1 1 rg 0 0 " + Format(width) + " " + Format(height) + " re f\n" +
            "0.75 0.75 0.75 RG 0.5 0.5 " + Format(boxWidth) + " " + Format(boxHeight) + " re S\n";

        if (selected) {
            double markLeft = Math.Max(2D, width * 0.2D);
            double markMidX = Math.Max(markLeft + 1D, width * 0.42D);
            double markRight = Math.Max(markMidX + 1D, width * 0.8D);
            double markMidY = Math.Max(2D, height * 0.25D);
            double markLeftY = Math.Min(height - 2D, height * 0.52D);
            double markRightY = Math.Min(height - 2D, height * 0.78D);
            content +=
                "0 0 0 RG 1.25 w " +
                Format(markLeft) + " " + Format(markLeftY) + " m " +
                Format(markMidX) + " " + Format(markMidY) + " l " +
                Format(markRight) + " " + Format(markRightY) + " l S\n";
        }

        return content + "Q\n";
    }

    private static string Format(double value) => value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
