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

        sb.Append(" ] /NeedAppearances false /DR << /Font << /Helv ")
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

    internal static string BuildTextFieldAppearanceContent(double width, double height, string value, double fontSize, PdfFormFieldStyle? style = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(value, nameof(value));
        Guard.Positive(fontSize, nameof(fontSize));

        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
        double textX = 3D;
        double textWidth = Math.Max(0D, width - 6D);
        string clippedValue = value;
        if (textWidth <= 0.001D) {
            clippedValue = string.Empty;
        }

        string content = "q\n";
        if (effectiveStyle.BackgroundColor.HasValue) {
            content += FormatColor(effectiveStyle.BackgroundColor.Value) + " rg 0 0 " + Format(width) + " " + Format(height) + " re f\n";
        }

        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0) {
            double inset = Math.Max(0.5D, effectiveStyle.BorderWidth * 0.5D);
            content += FormatColor(effectiveStyle.BorderColor.Value) + " RG " + Format(effectiveStyle.BorderWidth) + " w " +
                Format(inset) + " " + Format(inset) + " " + Format(Math.Max(0D, width - inset * 2D)) + " " + Format(Math.Max(0D, height - inset * 2D)) + " re S\n";
        }

        content += "BT /Helv " + Format(fontSize) + " Tf " + FormatColor(effectiveStyle.TextColor) + " rg " + Format(textX) + " " + Format(baseline) + " Td " + PdfSyntaxEscaper.WinAnsiHexString(clippedValue) + " Tj ET\n";
        return content + "Q\n";
    }

    internal static string BuildCheckBoxAppearanceContent(double width, double height, bool selected, PdfFormFieldStyle? style = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        string content = "q\n";
        if (effectiveStyle.BackgroundColor.HasValue) {
            content += FormatColor(effectiveStyle.BackgroundColor.Value) + " rg 0 0 " + Format(width) + " " + Format(height) + " re f\n";
        }

        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0) {
            double inset = Math.Max(0.5D, effectiveStyle.BorderWidth * 0.5D);
            content += FormatColor(effectiveStyle.BorderColor.Value) + " RG " + Format(effectiveStyle.BorderWidth) + " w " +
                Format(inset) + " " + Format(inset) + " " + Format(Math.Max(0D, width - inset * 2D)) + " " + Format(Math.Max(0D, height - inset * 2D)) + " re S\n";
        }

        if (selected) {
            double markLeft = Math.Max(2D, width * 0.2D);
            double markMidX = Math.Max(markLeft + 1D, width * 0.42D);
            double markRight = Math.Max(markMidX + 1D, width * 0.8D);
            double markMidY = Math.Max(2D, height * 0.25D);
            double markLeftY = Math.Min(height - 2D, height * 0.52D);
            double markRightY = Math.Min(height - 2D, height * 0.78D);
            content +=
                FormatColor(effectiveStyle.MarkColor) + " RG 1.25 w " +
                Format(markLeft) + " " + Format(markLeftY) + " m " +
                Format(markMidX) + " " + Format(markMidY) + " l " +
                Format(markRight) + " " + Format(markRightY) + " l S\n";
        }

        return content + "Q\n";
    }

    internal static string BuildRadioButtonAppearanceContent(double width, double height, bool selected, PdfFormFieldStyle? style = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        double centerX = width * 0.5D;
        double centerY = height * 0.5D;
        double radius = Math.Max(0D, Math.Min(width, height) * 0.5D - 0.75D);
        double control = radius * 0.5522847498D;
        string content = "q\n";
        if (effectiveStyle.BackgroundColor.HasValue) {
            content += FormatColor(effectiveStyle.BackgroundColor.Value) + " rg 0 0 " + Format(width) + " " + Format(height) + " re f\n";
        }

        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0) {
            content += FormatColor(effectiveStyle.BorderColor.Value) + " RG " + Format(effectiveStyle.BorderWidth) + " w " +
                Format(centerX + radius) + " " + Format(centerY) + " m " +
                Format(centerX + radius) + " " + Format(centerY + control) + " " + Format(centerX + control) + " " + Format(centerY + radius) + " " + Format(centerX) + " " + Format(centerY + radius) + " c " +
                Format(centerX - control) + " " + Format(centerY + radius) + " " + Format(centerX - radius) + " " + Format(centerY + control) + " " + Format(centerX - radius) + " " + Format(centerY) + " c " +
                Format(centerX - radius) + " " + Format(centerY - control) + " " + Format(centerX - control) + " " + Format(centerY - radius) + " " + Format(centerX) + " " + Format(centerY - radius) + " c " +
                Format(centerX + control) + " " + Format(centerY - radius) + " " + Format(centerX + radius) + " " + Format(centerY - control) + " " + Format(centerX + radius) + " " + Format(centerY) + " c S\n";
        }

        if (selected) {
            double dotRadius = Math.Max(0D, radius * 0.45D);
            double dotControl = dotRadius * 0.5522847498D;
            content +=
                FormatColor(effectiveStyle.MarkColor) + " rg " +
                Format(centerX + dotRadius) + " " + Format(centerY) + " m " +
                Format(centerX + dotRadius) + " " + Format(centerY + dotControl) + " " + Format(centerX + dotControl) + " " + Format(centerY + dotRadius) + " " + Format(centerX) + " " + Format(centerY + dotRadius) + " c " +
                Format(centerX - dotControl) + " " + Format(centerY + dotRadius) + " " + Format(centerX - dotRadius) + " " + Format(centerY + dotControl) + " " + Format(centerX - dotRadius) + " " + Format(centerY) + " c " +
                Format(centerX - dotRadius) + " " + Format(centerY - dotControl) + " " + Format(centerX - dotControl) + " " + Format(centerY - dotRadius) + " " + Format(centerX) + " " + Format(centerY - dotRadius) + " c " +
                Format(centerX + dotControl) + " " + Format(centerY - dotRadius) + " " + Format(centerX + dotRadius) + " " + Format(centerY - dotControl) + " " + Format(centerX + dotRadius) + " " + Format(centerY) + " c f\n";
        }

        return content + "Q\n";
    }

    internal static string FormatColor(PdfColor color) =>
        Format(color.R) + " " + Format(color.G) + " " + Format(color.B);

    private static string Format(double value) => value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
