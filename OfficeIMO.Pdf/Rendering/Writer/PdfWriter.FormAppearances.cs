namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] FormTextFieldLineSeparators = { '\n' };

    private static string BuildFormFieldTextAppearanceContent(
        double width,
        double height,
        string value,
        double fontSize,
        PdfFormFieldStyle? style,
        PdfOptions formOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        out IReadOnlyList<(string Name, int Id)> fontResources) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(value, nameof(value));
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NotNull(formOptions, nameof(formOptions));
        Guard.NotNull(ensureFont, nameof(ensureFont));

        PdfFormFieldStyle effectiveStyle = style ?? new PdfFormFieldStyle();
        string displayValue = PdfAcroFormDictionaryBuilder.GetTextFieldAppearanceDisplayValue(value, effectiveStyle);
        PdfFormFieldTextAlignment alignment = ResolveFormFieldTextAppearanceAlignment(effectiveStyle, formOptions);
        double availableTextWidth = Math.Max(0D, width - 6D);
        var resources = new List<(string Name, int Id)>();
        var sb = new StringBuilder();
        sb.Append("q\n");
        if (effectiveStyle.BackgroundColor.HasValue) {
            sb.Append(FormatAppearanceColor(effectiveStyle.BackgroundColor.Value))
                .Append(" rg 0 0 ")
                .Append(FormatAppearanceNumber(width))
                .Append(' ')
                .Append(FormatAppearanceNumber(height))
                .Append(" re f\n");
        }

        if (effectiveStyle.BorderColor.HasValue && effectiveStyle.BorderWidth > 0D) {
            double inset = Math.Max(0.5D, effectiveStyle.BorderWidth * 0.5D);
            sb.Append(FormatAppearanceColor(effectiveStyle.BorderColor.Value))
                .Append(" RG ")
                .Append(FormatAppearanceNumber(effectiveStyle.BorderWidth))
                .Append(" w ")
                .Append(FormatAppearanceNumber(inset))
                .Append(' ')
                .Append(FormatAppearanceNumber(inset))
                .Append(' ')
                .Append(FormatAppearanceNumber(Math.Max(0D, width - inset * 2D)))
                .Append(' ')
                .Append(FormatAppearanceNumber(Math.Max(0D, height - inset * 2D)))
                .Append(" re S\n");
        }

        if (effectiveStyle.IsMultiline) {
            AppendMultilineFormFieldTextAppearance(
                sb,
                height,
                availableTextWidth,
                displayValue,
                fontSize,
                effectiveStyle,
                alignment,
                formOptions,
                ensureFont,
                resources);
        } else if (effectiveStyle.IsComb && effectiveStyle.MaxLength.HasValue) {
            AppendCombFormFieldTextAppearance(
                sb,
                width,
                height,
                displayValue,
                fontSize,
                effectiveStyle,
                formOptions,
                ensureFont,
                resources);
        } else {
            double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
            string line = availableTextWidth <= 0.001D ? string.Empty : displayValue;
            AppendFormFieldTextAppearanceLine(
                sb,
                line,
                availableTextWidth,
                fontSize,
                baseline,
                effectiveStyle,
                alignment,
                formOptions,
                ensureFont,
                resources);
        }

        sb.Append("Q\n");
        fontResources = resources;
        return sb.ToString();
    }

    private static void AppendMultilineFormFieldTextAppearance(
        StringBuilder sb,
        double height,
        double availableTextWidth,
        string displayValue,
        double fontSize,
        PdfFormFieldStyle effectiveStyle,
        PdfFormFieldTextAlignment alignment,
        PdfOptions formOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        List<(string Name, int Id)> fontResources) {
        string[] lines = SplitFormTextFieldAppearanceLines(displayValue);
        double lineHeight = Math.Max(fontSize, fontSize * 1.2D);
        double baseline = Math.Max(2D, height - fontSize * 1.15D);
        for (int i = 0; i < lines.Length && baseline >= 2D; i++) {
            string line = availableTextWidth <= 0.001D ? string.Empty : lines[i];
            AppendFormFieldTextAppearanceLine(
                sb,
                line,
                availableTextWidth,
                fontSize,
                baseline,
                effectiveStyle,
                alignment,
                formOptions,
                ensureFont,
                fontResources);
            baseline -= lineHeight;
        }
    }

    private static void AppendCombFormFieldTextAppearance(
        StringBuilder sb,
        double width,
        double height,
        string displayValue,
        double fontSize,
        PdfFormFieldStyle effectiveStyle,
        PdfOptions formOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        List<(string Name, int Id)> fontResources) {
        int cellCount = effectiveStyle.MaxLength!.Value;
        double cellWidth = width / cellCount;
        double baseline = Math.Max(2D, (height - fontSize) / 2D + fontSize * 0.72D);
        int glyphIndex = 0;
        for (int valueIndex = 0; valueIndex < displayValue.Length && glyphIndex < cellCount; glyphIndex++) {
            int scalarLength = GetScalarUtf16Length(displayValue, valueIndex);
            string glyph = displayValue.Substring(valueIndex, scalarLength);
            valueIndex += scalarLength;
            IReadOnlyList<RichSeg> segments = BuildFormAppearanceSegments(glyph, fontSize, effectiveStyle, formOptions);
            EnsureCanWriteFormAppearanceLine(segments, formOptions);
            double glyphWidth = MeasureRichLineWidth(segments, formOptions);
            double textX = glyphIndex * cellWidth + Math.Max(0D, (cellWidth - glyphWidth) / 2D);
            AppendFormAppearanceSegments(
                sb,
                segments,
                textX,
                baseline,
                effectiveStyle.TextColor,
                formOptions,
                ensureFont,
                fontResources);
        }
    }

    private static void AppendFormFieldTextAppearanceLine(
        StringBuilder sb,
        string line,
        double availableTextWidth,
        double fontSize,
        double baseline,
        PdfFormFieldStyle effectiveStyle,
        PdfFormFieldTextAlignment alignment,
        PdfOptions formOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        List<(string Name, int Id)> fontResources) {
        IReadOnlyList<RichSeg> segments = BuildFormAppearanceSegments(line, fontSize, effectiveStyle, formOptions);
        EnsureCanWriteFormAppearanceLine(segments, formOptions);
        double lineWidth = MeasureRichLineWidth(segments, formOptions);
        double textX = CalculateFormFieldAppearanceAlignedTextX(availableTextWidth, lineWidth, alignment);
        AppendFormAppearanceSegments(
            sb,
            segments,
            textX,
            baseline,
            effectiveStyle.TextColor,
            formOptions,
            ensureFont,
            fontResources);
    }

    private static IReadOnlyList<RichSeg> BuildFormAppearanceSegments(string text, double fontSize, PdfFormFieldStyle effectiveStyle, PdfOptions formOptions) {
        var runs = new[] {
            TextRun.Normal(text, effectiveStyle.TextColor, fontSize, font: PdfStandardFont.Helvetica)
        };
        var wrapped = WrapRichRunsCore(
            runs,
            Math.Max(1D, 1000000D),
            fontSize,
            PdfStandardFont.Helvetica,
            Math.Max(fontSize, fontSize * 1.2D),
            firstLineWidthPts: null,
            DefaultParagraphTabStopWidth,
            formOptions);
        return wrapped.Lines.Count == 0
            ? Array.Empty<RichSeg>()
            : wrapped.Lines[0];
    }

    private static void EnsureCanWriteFormAppearanceLine(IReadOnlyList<RichSeg> line, PdfOptions formOptions) {
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (segment.LeadingSpace && !CanEncodeAppearanceText(" ", segment.Font, formOptions)) {
                throw CreateUnsupportedFormAppearanceTextException(" ");
            }

            if (!CanEncodeAppearanceText(segment.Text, segment.Font, formOptions)) {
                throw CreateUnsupportedFormAppearanceTextException(segment.Text);
            }
        }
    }

    private static InvalidOperationException CreateUnsupportedFormAppearanceTextException(string text) =>
        new InvalidOperationException("PDF form field appearance text contains glyphs that cannot be written by the selected standard or embedded fallback fonts: " + text);

    private static void AppendFormAppearanceSegments(
        StringBuilder sb,
        IReadOnlyList<RichSeg> line,
        double x,
        double y,
        PdfColor defaultTextColor,
        PdfOptions formOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        List<(string Name, int Id)> fontResources) {
        double xCursor = 0D;
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            string fontResource = GetFreeTextAppearanceFontResourceName(segment.Font, formOptions);
            RegisterAppearanceFontResource(fontResources, fontResource, ensureFont(segment.Font, formOptions));
            double runFontSize = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
            double textRise = TextRiseForBaseline(segment.FontSize, segment.Baseline);
            if (segment.LeadingSpace) {
                double gap = segment.LeadingAdvance > 0D
                    ? segment.LeadingAdvance
                    : MeasureRichText(" ", segment.Font, segment.FontSize, segment.Baseline, formOptions);
                AppendFreeTextAppearanceTextSegment(
                    sb,
                    fontResource,
                    runFontSize,
                    textRise,
                    x + xCursor,
                    y,
                    " ",
                    segment.Font,
                    segment.Color ?? defaultTextColor,
                    formOptions);
                xCursor += gap;
            }

            if (segment.Text.Length > 0) {
                AppendFreeTextAppearanceTextSegment(
                    sb,
                    fontResource,
                    runFontSize,
                    textRise,
                    x + xCursor,
                    y,
                    segment.Text,
                    segment.Font,
                    segment.Color ?? defaultTextColor,
                    formOptions);
                xCursor += MeasureRichText(segment.Text, segment.Font, segment.FontSize, segment.Baseline, formOptions);
            }
        }
    }

    private static PdfFormFieldTextAlignment ResolveFormFieldTextAppearanceAlignment(PdfFormFieldStyle effectiveStyle, PdfOptions formOptions) =>
        effectiveStyle.TextAlignment ?? formOptions.AcroFormDefaultTextAlignmentSnapshot ?? PdfFormFieldTextAlignment.Left;

    private static double CalculateFormFieldAppearanceAlignedTextX(double availableTextWidth, double measuredTextWidth, PdfFormFieldTextAlignment alignment) {
        double padding = 3D;
        if (availableTextWidth <= 0D || measuredTextWidth >= availableTextWidth) {
            return padding;
        }

        return alignment switch {
            PdfFormFieldTextAlignment.Center => padding + (availableTextWidth - measuredTextWidth) / 2D,
            PdfFormFieldTextAlignment.Right => padding + availableTextWidth - measuredTextWidth,
            _ => padding
        };
    }

    private static string[] SplitFormTextFieldAppearanceLines(string value) {
        return value
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Split(FormTextFieldLineSeparators);
    }
}
