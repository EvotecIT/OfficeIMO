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
        IReadOnlyList<(string Name, int Id)> fontResources = helveticaFontId > 0
            ? new[] { ("Helv", helveticaFontId) }
            : Array.Empty<(string Name, int Id)>();

        return BuildTextFieldAppearanceStreamDictionary(width, height, fontResources, contentLength);
    }

    internal static string BuildTextFieldAppearanceStreamDictionary(double width, double height, IReadOnlyList<(string Name, int Id)> fontResources, int contentLength) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(fontResources, nameof(fontResources));
        if (contentLength < 0) {
            throw new ArgumentOutOfRangeException(nameof(contentLength), "PDF appearance stream length cannot be negative.");
        }

        var sb = new StringBuilder();
        sb.Append("<< /Type /XObject /Subtype /Form /BBox [0 0 ")
            .Append(Format(width))
            .Append(' ')
            .Append(Format(height))
            .Append(']');
        if (fontResources.Count > 0) {
            sb.Append(" /Resources << /Font <<");
            for (int i = 0; i < fontResources.Count; i++) {
                (string name, int id) = fontResources[i];
                Guard.NotNullOrWhiteSpace(name, nameof(fontResources));
                if (id <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(fontResources), id, "PDF appearance font resource object id must be positive.");
                }

                string normalizedName = name[0] == '/' ? name.Substring(1) : name;
                Guard.NotNullOrWhiteSpace(normalizedName, nameof(fontResources));
                sb.Append(" /")
                    .Append(PdfSyntaxEscaper.Name(normalizedName))
                    .Append(' ')
                    .Append(PdfSyntaxEscaper.IndirectReference(id));
            }

            sb.Append(" >> >>");
        }

        sb.Append(" /Length ")
            .Append(contentLength.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(" >>");
        return sb.ToString();
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

    internal static string BuildTextFieldAppearanceContent(double width, double height, string value, double fontSize, PdfFormFieldStyle? style = null, string? encodedTextHex = null, PdfFormFieldTextAlignment? textAlignment = null, double? textWidth = null, string? fontResourceName = null, Func<string, string?>? encodeTextSegmentHex = null, Func<string, double, double>? measureTextSegmentWidth = null, Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? encodeTextSegments = null) {
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
        string effectiveFontResourceName = PdfSyntaxEscaper.Name(string.IsNullOrWhiteSpace(fontResourceName) ? "Helv" : fontResourceName!);

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
            content += BuildMultilineTextFieldAppearanceContent(height, displayValue, fontSize, effectiveStyle, effectiveAlignment, availableTextWidth, effectiveFontResourceName, encodeTextSegmentHex, measureTextSegmentWidth, encodeTextSegments);
        } else if (effectiveStyle.IsComb && effectiveStyle.MaxLength.HasValue) {
            content += BuildCombTextFieldAppearanceContent(width, height, displayValue, fontSize, effectiveStyle, effectiveFontResourceName, encodeTextSegmentHex, measureTextSegmentWidth, encodeTextSegments);
        } else {
            double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
            double measuredTextWidth = textWidth.HasValue ? textWidth.Value : MeasureTextAppearanceSegment(displayValue, fontSize, measureTextSegmentWidth);
            double textX = CalculateAlignedTextX(availableTextWidth, measuredTextWidth, effectiveAlignment);
            string clippedValue = availableTextWidth <= 0.001D ? string.Empty : displayValue;
            string textShowing = BuildTextAppearanceShowing(clippedValue, fontSize, encodedTextHex, encodeTextSegmentHex, encodeTextSegments);
            content += "BT /" + effectiveFontResourceName + " " + Format(fontSize) + " Tf " + FormatColor(effectiveStyle.TextColor) + " rg " + Format(textX) + " " + Format(baseline) + " Td " + textShowing + " ET\n";
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

    private static string BuildMultilineTextFieldAppearanceContent(double height, string displayValue, double fontSize, PdfFormFieldStyle effectiveStyle, PdfFormFieldTextAlignment alignment, double availableTextWidth, string fontResourceName, Func<string, string?>? encodeTextSegmentHex, Func<string, double, double>? measureTextSegmentWidth, Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? encodeTextSegments) {
        string[] lines = SplitTextFieldAppearanceLines(displayValue);
        double lineHeight = Math.Max(fontSize, fontSize * 1.2D);
        double baseline = Math.Max(2D, height - fontSize * 1.15D);
        string content = string.Empty;
        for (int i = 0; i < lines.Length && baseline >= 2D; i++) {
            string line = availableTextWidth <= 0.001D ? string.Empty : lines[i];
            double lineWidth = MeasureTextAppearanceSegment(line, fontSize, measureTextSegmentWidth);
            double textX = CalculateAlignedTextX(availableTextWidth, lineWidth, alignment);
            string textShowing = BuildTextAppearanceShowing(line, fontSize, encodedTextHex: null, encodeTextSegmentHex, encodeTextSegments);
            content += "BT /" + fontResourceName + " " + Format(fontSize) + " Tf " + FormatColor(effectiveStyle.TextColor) + " rg " + Format(textX) + " " + Format(baseline) + " Td " + textShowing + " ET\n";
            baseline -= lineHeight;
        }

        return content;
    }

    private static string BuildCombTextFieldAppearanceContent(double width, double height, string displayValue, double fontSize, PdfFormFieldStyle effectiveStyle, string fontResourceName, Func<string, string?>? encodeTextSegmentHex, Func<string, double, double>? measureTextSegmentWidth, Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? encodeTextSegments) {
        int cellCount = effectiveStyle.MaxLength!.Value;
        double cellWidth = width / cellCount;
        double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
        var content = new StringBuilder();
        int glyphIndex = 0;
        for (int valueIndex = 0; valueIndex < displayValue.Length && glyphIndex < cellCount; glyphIndex++) {
            int scalarLength = GetScalarLength(displayValue, valueIndex);
            string glyph = displayValue.Substring(valueIndex, scalarLength);
            valueIndex += scalarLength;
            double glyphWidth = MeasureTextAppearanceSegment(glyph, fontSize, measureTextSegmentWidth);
            double textX = glyphIndex * cellWidth + Math.Max(0D, (cellWidth - glyphWidth) / 2D);
            string textShowing = BuildTextAppearanceShowing(glyph, fontSize, encodedTextHex: null, encodeTextSegmentHex, encodeTextSegments);
            content.Append("BT /")
                .Append(fontResourceName)
                .Append(' ')
                .Append(Format(fontSize))
                .Append(" Tf ")
                .Append(FormatColor(effectiveStyle.TextColor))
                .Append(" rg ")
                .Append(Format(textX))
                .Append(' ')
                .Append(Format(baseline))
                .Append(" Td ")
                .Append(textShowing)
                .Append(" ET\n");
        }

        return content.ToString();
    }

    private static string BuildTextAppearanceShowing(string value, double fontSize, string? encodedTextHex, Func<string, string?>? encodeTextSegmentHex, Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? encodeTextSegments) {
        if (value.Length > 0 && encodeTextSegments != null) {
            IReadOnlyList<PdfTextAppearanceSegment> segments = encodeTextSegments(value);
            if (segments.Count > 0) {
                var sb = new StringBuilder();
                for (int i = 0; i < segments.Count; i++) {
                    PdfTextAppearanceSegment segment = segments[i];
                    if (segment.EncodedHex.Length == 0) {
                        continue;
                    }

                    if (sb.Length > 0) {
                        sb.Append(' ');
                    }

                    sb.Append('/')
                        .Append(PdfSyntaxEscaper.Name(segment.FontResourceName))
                        .Append(' ')
                        .Append(Format(fontSize))
                        .Append(" Tf <")
                        .Append(segment.EncodedHex)
                        .Append("> Tj");
                }

                if (sb.Length > 0) {
                    return sb.ToString();
                }
            }
        }

        return EncodeTextAppearanceSegment(value, encodedTextHex, encodeTextSegmentHex) + " Tj";
    }

    private static string EncodeTextAppearanceSegment(string value, string? encodedTextHex, Func<string, string?>? encodeTextSegmentHex) {
        if (value.Length == 0) {
            return PdfSyntaxEscaper.WinAnsiHexString(value);
        }

        if (!string.IsNullOrEmpty(encodedTextHex)) {
            return "<" + encodedTextHex + ">";
        }

        string? segmentHex = encodeTextSegmentHex?.Invoke(value);
        return string.IsNullOrEmpty(segmentHex)
            ? PdfSyntaxEscaper.WinAnsiHexString(value)
            : "<" + segmentHex + ">";
    }

    private static double MeasureTextAppearanceSegment(string value, double fontSize, Func<string, double, double>? measureTextSegmentWidth) =>
        measureTextSegmentWidth == null
            ? PdfWriter.EstimateSimpleTextWidth(value, PdfStandardFont.Helvetica, fontSize)
            : measureTextSegmentWidth(value, fontSize);

    private static int GetScalarLength(string value, int index) =>
        char.IsHighSurrogate(value[index]) &&
        index + 1 < value.Length &&
        char.IsLowSurrogate(value[index + 1])
            ? 2
            : 1;

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
