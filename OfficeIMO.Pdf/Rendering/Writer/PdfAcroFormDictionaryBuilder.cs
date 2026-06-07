namespace OfficeIMO.Pdf;

internal static class PdfAcroFormDictionaryBuilder {
    private static readonly char[] TextFieldLineSeparators = { '\n' };

    internal static string BuildAcroFormDictionary(IReadOnlyList<int> fieldObjectIds, int helveticaFontId, PdfFormFieldTextAlignment? defaultTextAlignment = null) {
        Guard.NotNull(fieldObjectIds, nameof(fieldObjectIds));
        if (fieldObjectIds.Count == 0) {
            throw new ArgumentException("PDF AcroForm dictionary requires at least one field object.", nameof(fieldObjectIds));
        }

        if (defaultTextAlignment.HasValue) {
            Guard.FormFieldTextAlignment(defaultTextAlignment.Value, nameof(defaultTextAlignment));
        }

        var sb = new StringBuilder();
        sb.Append("<< /Fields [");
        for (int i = 0; i < fieldObjectIds.Count; i++) {
            sb.Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(fieldObjectIds[i]));
        }

        sb.Append(" ] /NeedAppearances false /DR << /Font << /Helv ")
            .Append(PdfSyntaxEscaper.IndirectReference(helveticaFontId))
            .Append(" >> >> /DA (/Helv 10 Tf 0 g)");
        if (defaultTextAlignment.HasValue) {
            sb.Append(" /Q ")
                .Append(ToQuadding(defaultTextAlignment.Value));
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static int ToQuadding(PdfFormFieldTextAlignment alignment) {
        switch (alignment) {
            case PdfFormFieldTextAlignment.Left:
                return 0;
            case PdfFormFieldTextAlignment.Center:
                return 1;
            case PdfFormFieldTextAlignment.Right:
                return 2;
            default:
                throw new ArgumentOutOfRangeException(nameof(alignment), "PDF form field text alignment must be Left, Center, or Right.");
        }
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

    internal static string BuildTextFieldAppearanceContent(double width, double height, string value, double fontSize, PdfFormFieldStyle? style = null, string? encodedTextHex = null, PdfFormFieldTextAlignment? textAlignment = null, double? textWidth = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(value, nameof(value));
        Guard.Positive(fontSize, nameof(fontSize));
        if (textAlignment.HasValue) {
            Guard.FormFieldTextAlignment(textAlignment.Value, nameof(textAlignment));
        }

        if (textWidth.HasValue && (textWidth.Value < 0 || double.IsNaN(textWidth.Value) || double.IsInfinity(textWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textWidth), textWidth.Value, "PDF appearance text width must be a finite non-negative number.");
        }

        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        string displayValue = GetTextFieldAppearanceDisplayValue(value, effectiveStyle);
        PdfFormFieldTextAlignment effectiveAlignment = textAlignment ?? effectiveStyle.TextAlignment ?? PdfFormFieldTextAlignment.Left;
        double availableTextWidth = Math.Max(0D, width - 6D);

        string content = "q\n";
        if (effectiveStyle.BackgroundColor.HasValue) {
            content += FormatColor(effectiveStyle.BackgroundColor.Value) + " rg 0 0 " + Format(width) + " " + Format(height) + " re f\n";
        }

        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0) {
            double inset = Math.Max(0.5D, effectiveStyle.BorderWidth * 0.5D);
            content += FormatColor(effectiveStyle.BorderColor.Value) + " RG " + Format(effectiveStyle.BorderWidth) + " w " +
                Format(inset) + " " + Format(inset) + " " + Format(Math.Max(0D, width - inset * 2D)) + " " + Format(Math.Max(0D, height - inset * 2D)) + " re S\n";
        }

        if (effectiveStyle.IsMultiline) {
            content += BuildMultilineTextFieldAppearanceContent(height, displayValue, fontSize, effectiveStyle, effectiveAlignment, availableTextWidth);
        } else {
            double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
            double measuredTextWidth = textWidth.HasValue ? textWidth.Value : PdfWriter.EstimateSimpleTextWidth(displayValue, PdfStandardFont.Helvetica, fontSize);
            double textX = CalculateAlignedTextX(availableTextWidth, measuredTextWidth, effectiveAlignment);
            string clippedValue = availableTextWidth <= 0.001D ? string.Empty : displayValue;
            string textHex = encodedTextHex == null || clippedValue.Length == 0 ? PdfSyntaxEscaper.WinAnsiHexString(clippedValue) : "<" + encodedTextHex + ">";
            content += "BT /Helv " + Format(fontSize) + " Tf " + FormatColor(effectiveStyle.TextColor) + " rg " + Format(textX) + " " + Format(baseline) + " Td " + textHex + " Tj ET\n";
        }

        return content + "Q\n";
    }

    internal static string GetTextFieldAppearanceDisplayValue(string value, PdfFormFieldStyle? style) {
        Guard.NotNull(value, nameof(value));
        return style != null && style.IsPassword
            ? new string('*', value.Length)
            : value;
    }

    private static double CalculateAlignedTextX(double availableTextWidth, double measuredTextWidth, PdfFormFieldTextAlignment alignment) {
        double padding = 3D;
        if (availableTextWidth <= 0D || measuredTextWidth >= availableTextWidth) {
            return padding;
        }

        switch (alignment) {
            case PdfFormFieldTextAlignment.Left:
                return padding;
            case PdfFormFieldTextAlignment.Center:
                return padding + (availableTextWidth - measuredTextWidth) / 2D;
            case PdfFormFieldTextAlignment.Right:
                return padding + availableTextWidth - measuredTextWidth;
            default:
                throw new ArgumentOutOfRangeException(nameof(alignment), "PDF form field text alignment must be Left, Center, or Right.");
        }
    }

    private static string BuildMultilineTextFieldAppearanceContent(double height, string displayValue, double fontSize, PdfFormFieldStyle effectiveStyle, PdfFormFieldTextAlignment alignment, double availableTextWidth) {
        string[] lines = SplitTextFieldAppearanceLines(displayValue);
        double lineHeight = Math.Max(fontSize, fontSize * 1.2D);
        double baseline = Math.Max(2D, height - fontSize * 1.15D);
        string content = string.Empty;
        for (int i = 0; i < lines.Length && baseline >= 2D; i++) {
            string line = availableTextWidth <= 0.001D ? string.Empty : lines[i];
            double lineWidth = PdfWriter.EstimateSimpleTextWidth(line, PdfStandardFont.Helvetica, fontSize);
            double textX = CalculateAlignedTextX(availableTextWidth, lineWidth, alignment);
            string textHex = PdfSyntaxEscaper.WinAnsiHexString(line);
            content += "BT /Helv " + Format(fontSize) + " Tf " + FormatColor(effectiveStyle.TextColor) + " rg " + Format(textX) + " " + Format(baseline) + " Td " + textHex + " Tj ET\n";
            baseline -= lineHeight;
        }

        return content;
    }

    private static string[] SplitTextFieldAppearanceLines(string value) {
        return value
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Split(TextFieldLineSeparators);
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
