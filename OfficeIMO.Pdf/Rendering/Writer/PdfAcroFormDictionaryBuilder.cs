namespace OfficeIMO.Pdf;

internal static class PdfAcroFormDictionaryBuilder {
    private static readonly double[] DefaultBorderDashPattern = { 3D };

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
            content += BuildRectangularBorderAppearanceContent(width, height, effectiveStyle);
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
        List<string> lines = WrapTextFieldAppearanceLines(displayValue, fontSize, availableTextWidth, measureTextSegmentWidth);
        double lineHeight = Math.Max(fontSize, fontSize * 1.2D);
        double baseline = Math.Max(2D, height - fontSize * 1.15D);
        var content = new StringBuilder();
        for (int i = 0; i < lines.Count && baseline >= 2D; i++) {
            string line = availableTextWidth <= 0.001D ? string.Empty : lines[i];
            double lineWidth = MeasureTextAppearanceSegment(line, fontSize, measureTextSegmentWidth);
            double textX = CalculateAlignedTextX(availableTextWidth, lineWidth, alignment);
            string textShowing = BuildTextAppearanceShowing(line, fontSize, encodedTextHex: null, encodeTextSegmentHex, encodeTextSegments);
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
            baseline -= lineHeight;
        }

        return content.ToString();
    }

    private static List<string> WrapTextFieldAppearanceLines(string displayValue, double fontSize, double availableTextWidth, Func<string, double, double>? measureTextSegmentWidth) {
        string[] explicitLines = SplitTextFieldAppearanceLines(displayValue);
        var wrappedLines = new List<string>();
        for (int i = 0; i < explicitLines.Length; i++) {
            WrapTextFieldAppearanceLine(explicitLines[i], fontSize, availableTextWidth, measureTextSegmentWidth, wrappedLines);
        }

        return wrappedLines;
    }

    private static void WrapTextFieldAppearanceLine(string line, double fontSize, double availableTextWidth, Func<string, double, double>? measureTextSegmentWidth, List<string> wrappedLines) {
        if (line.Length == 0 || availableTextWidth <= 0.001D) {
            wrappedLines.Add(line);
            return;
        }

        string currentLine = string.Empty;
        foreach (string token in EnumerateTextFieldAppearanceTokens(line)) {
            if (currentLine.Length == 0 && IsAllWhiteSpace(token)) {
                continue;
            }

            string candidate = currentLine + token;
            if (FitsTextFieldAppearanceLine(candidate, fontSize, availableTextWidth, measureTextSegmentWidth)) {
                currentLine = candidate;
                continue;
            }

            if (IsAllWhiteSpace(token)) {
                AddWrappedTextFieldAppearanceLine(wrappedLines, currentLine);
                currentLine = string.Empty;
                continue;
            }

            if (currentLine.Length > 0) {
                AddWrappedTextFieldAppearanceLine(wrappedLines, currentLine);
                currentLine = string.Empty;
            }

            currentLine = AddWrappedTextFieldAppearanceWord(token, fontSize, availableTextWidth, measureTextSegmentWidth, wrappedLines);
        }

        AddWrappedTextFieldAppearanceLine(wrappedLines, currentLine);
    }

    private static IEnumerable<string> EnumerateTextFieldAppearanceTokens(string line) {
        int start = 0;
        bool inWhiteSpace = char.IsWhiteSpace(line[0]);
        for (int i = 1; i < line.Length; i++) {
            bool isWhiteSpace = char.IsWhiteSpace(line[i]);
            if (isWhiteSpace == inWhiteSpace) {
                continue;
            }

            yield return line.Substring(start, i - start);
            start = i;
            inWhiteSpace = isWhiteSpace;
        }

        yield return line.Substring(start);
    }

    private static string AddWrappedTextFieldAppearanceWord(string word, double fontSize, double availableTextWidth, Func<string, double, double>? measureTextSegmentWidth, List<string> wrappedLines) {
        string currentLine = string.Empty;
        for (int i = 0; i < word.Length;) {
            int scalarLength = GetScalarLength(word, i);
            string scalar = word.Substring(i, scalarLength);
            i += scalarLength;

            string candidate = currentLine + scalar;
            if (currentLine.Length == 0 || FitsTextFieldAppearanceLine(candidate, fontSize, availableTextWidth, measureTextSegmentWidth)) {
                currentLine = candidate;
                continue;
            }

            AddWrappedTextFieldAppearanceLine(wrappedLines, currentLine);
            currentLine = scalar;
        }

        return currentLine;
    }

    private static void AddWrappedTextFieldAppearanceLine(List<string> wrappedLines, string line) {
        if (line.Length == 0) {
            wrappedLines.Add(string.Empty);
            return;
        }

        wrappedLines.Add(line.TrimEnd());
    }

    private static bool FitsTextFieldAppearanceLine(string line, double fontSize, double availableTextWidth, Func<string, double, double>? measureTextSegmentWidth) =>
        MeasureTextAppearanceSegment(line.TrimEnd(), fontSize, measureTextSegmentWidth) <= availableTextWidth;

    private static bool IsAllWhiteSpace(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) {
                return false;
            }
        }

        return true;
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

        if (encodeTextSegmentHex != null) {
            string? segmentHex = encodeTextSegmentHex(value);
            if (string.IsNullOrEmpty(segmentHex)) {
                throw new ArgumentException("Text field appearance segment cannot be encoded by the selected embedded appearance font.", nameof(value));
            }

            return "<" + segmentHex + ">";
        }

        return PdfSyntaxEscaper.WinAnsiHexString(value);
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
            content += BuildRectangularBorderAppearanceContent(width, height, effectiveStyle);
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
            if (effectiveStyle.BorderStyle == PdfFormFieldBorderStyle.Underline) {
                content += BuildUnderlineBorderAppearanceContent(width, effectiveStyle);
            } else {
                content += BuildBorderStrokeOperators(effectiveStyle.BorderColor.Value, effectiveStyle.BorderWidth, GetEffectiveBorderDashPattern(effectiveStyle)) +
                    Format(centerX + radius) + " " + Format(centerY) + " m " +
                    Format(centerX + radius) + " " + Format(centerY + control) + " " + Format(centerX + control) + " " + Format(centerY + radius) + " " + Format(centerX) + " " + Format(centerY + radius) + " c " +
                    Format(centerX - control) + " " + Format(centerY + radius) + " " + Format(centerX - radius) + " " + Format(centerY + control) + " " + Format(centerX - radius) + " " + Format(centerY) + " c " +
                    Format(centerX - radius) + " " + Format(centerY - control) + " " + Format(centerX - control) + " " + Format(centerY - radius) + " " + Format(centerX) + " " + Format(centerY - radius) + " c " +
                    Format(centerX + control) + " " + Format(centerY - radius) + " " + Format(centerX + radius) + " " + Format(centerY - control) + " " + Format(centerX + radius) + " " + Format(centerY) + " c S\n";
            }
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

    private static string BuildRectangularBorderAppearanceContent(double width, double height, PdfFormFieldStyle style) {
        if (style.BorderColor == null || style.BorderWidth <= 0D) {
            return string.Empty;
        }

        return BuildRectangularBorderAppearanceContent(width, height, style.BorderColor.Value, style.BorderWidth, GetEffectiveBorderDashPattern(style), style.BorderStyle);
    }

    internal static string BuildRectangularBorderAppearanceContent(double width, double height, PdfColor borderColor, double borderWidth, IReadOnlyList<double>? dashPattern, PdfFormFieldBorderStyle borderStyle) {
        if (borderWidth <= 0D) {
            return string.Empty;
        }

        if (borderStyle == PdfFormFieldBorderStyle.Underline) {
            return BuildUnderlineBorderAppearanceContent(width, borderColor, borderWidth);
        }

        if (borderStyle == PdfFormFieldBorderStyle.Beveled || borderStyle == PdfFormFieldBorderStyle.Inset) {
            return BuildBeveledOrInsetBorderAppearanceContent(width, height, borderColor, borderWidth, borderStyle);
        }

        double inset = Math.Max(0.5D, borderWidth * 0.5D);
        return BuildBorderStrokeOperators(borderColor, borderWidth, dashPattern) +
            Format(inset) + " " + Format(inset) + " " + Format(Math.Max(0D, width - inset * 2D)) + " " + Format(Math.Max(0D, height - inset * 2D)) + " re S\n";
    }

    private static string BuildUnderlineBorderAppearanceContent(double width, PdfFormFieldStyle style) {
        if (style.BorderColor == null || style.BorderWidth <= 0D) {
            return string.Empty;
        }

        return BuildUnderlineBorderAppearanceContent(width, style.BorderColor.Value, style.BorderWidth);
    }

    private static string BuildUnderlineBorderAppearanceContent(double width, PdfColor borderColor, double borderWidth) {
        double y = Math.Max(0.5D, borderWidth * 0.5D);
        return BuildBorderStrokeOperators(borderColor, borderWidth, null) +
            "0 " + Format(y) + " m " + Format(width) + " " + Format(y) + " l S\n";
    }

    private static string BuildBeveledOrInsetBorderAppearanceContent(double width, double height, PdfColor borderColor, double borderWidth, PdfFormFieldBorderStyle borderStyle) {
        double inset = Math.Max(0.5D, borderWidth * 0.5D);
        double right = Math.Max(inset, width - inset);
        double top = Math.Max(inset, height - inset);
        PdfColor light = Lighten(borderColor);
        PdfColor dark = Darken(borderColor);
        PdfColor topLeft = borderStyle == PdfFormFieldBorderStyle.Inset ? dark : light;
        PdfColor bottomRight = borderStyle == PdfFormFieldBorderStyle.Inset ? light : dark;

        return BuildBorderStrokeOperators(topLeft, borderWidth, null) +
            Format(inset) + " " + Format(inset) + " m " +
            Format(inset) + " " + Format(top) + " l " +
            Format(right) + " " + Format(top) + " l S\n" +
            BuildBorderStrokeOperators(bottomRight, borderWidth, null) +
            Format(inset) + " " + Format(inset) + " m " +
            Format(right) + " " + Format(inset) + " l " +
            Format(right) + " " + Format(top) + " l S\n";
    }

    private static PdfColor Lighten(PdfColor color) => new PdfColor(Lighten(color.R), Lighten(color.G), Lighten(color.B));

    private static double Lighten(double component) => component + (1D - component) * 0.55D;

    private static PdfColor Darken(PdfColor color) => new PdfColor(color.R * 0.45D, color.G * 0.45D, color.B * 0.45D);

    private static IReadOnlyList<double>? GetEffectiveBorderDashPattern(PdfFormFieldStyle style) {
        if (style.BorderDashPattern != null && style.BorderDashPattern.Count > 0) {
            return style.BorderDashPattern;
        }

        return style.BorderStyle == PdfFormFieldBorderStyle.Dashed
            ? DefaultBorderDashPattern
            : null;
    }

    internal static string BuildBorderStrokeOperators(PdfColor color, double borderWidth, IReadOnlyList<double>? dashPattern) {
        string operators = FormatColor(color) + " RG " + Format(borderWidth) + " w ";
        if (dashPattern == null || dashPattern.Count == 0) {
            return operators;
        }

        var builder = new StringBuilder(operators);
        builder.Append('[');
        for (int i = 0; i < dashPattern.Count; i++) {
            if (i > 0) {
                builder.Append(' ');
            }

            builder.Append(Format(dashPattern[i]));
        }

        builder.Append("] 0 d ");
        return builder.ToString();
    }

    private static string Format(double value) => value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
