using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFreeTextAnnotationAppearanceContent(
        FreeTextAnnotation annotation,
        double width,
        double height,
        PdfOptions pageOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        out IReadOnlyList<(string Name, int Id)> fontResources) {
        Guard.NotNull(annotation, nameof(annotation));
        Guard.NotNull(pageOptions, nameof(pageOptions));
        Guard.NotNull(ensureFont, nameof(ensureFont));

        PdfColor resolvedTextColor = annotation.TextColor;
        double effectiveLineHeight = annotation.LineHeight ?? annotation.FontSize * 1.2D;
        double availableWidth = Math.Max(0D, width - annotation.Padding * 2D);
        double availableHeight = Math.Max(0D, height - annotation.Padding * 2D);
        var runs = new[] {
            TextRun.Normal(annotation.Contents, resolvedTextColor, annotation.FontSize, font: PdfStandardFont.Helvetica)
        };
        var wrapped = WrapRichRunsCore(
            runs,
            availableWidth,
            annotation.FontSize,
            PdfStandardFont.Helvetica,
            effectiveLineHeight,
            firstLineWidthPts: null,
            DefaultParagraphTabStopWidth,
            pageOptions);

        int maxVisibleLines = availableHeight > 0D
            ? Math.Max(1, (int)Math.Floor(availableHeight / effectiveLineHeight))
            : 0;

        var resources = new List<(string Name, int Id)>();
        var sb = new StringBuilder();
        sb.Append("q\n");
        if (annotation.FillColor.HasValue) {
            sb.Append(FormatAppearanceColor(annotation.FillColor.Value))
                .Append(" rg 0 0 ")
                .Append(FormatAppearanceNumber(width))
                .Append(' ')
                .Append(FormatAppearanceNumber(height))
                .Append(" re f\n");
        }

        if (annotation.BorderColor.HasValue && annotation.BorderWidth > 0D) {
            double inset = Math.Max(0.5D, annotation.BorderWidth * 0.5D);
            sb.Append(FormatAppearanceColor(annotation.BorderColor.Value))
                .Append(" RG ")
                .Append(FormatAppearanceNumber(annotation.BorderWidth))
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

        double baseline = height - annotation.Padding - annotation.FontSize;
        int visibleLineCount = maxVisibleLines > 0
            ? Math.Min(maxVisibleLines, wrapped.Lines.Count)
            : 0;
        for (int lineIndex = 0; lineIndex < visibleLineCount && baseline >= annotation.Padding - annotation.FontSize * 0.25D; lineIndex++) {
            IReadOnlyList<RichSeg> line = wrapped.Lines[lineIndex];
            if (!CanEncodeFreeTextAppearanceLine(line, pageOptions)) {
                baseline -= effectiveLineHeight;
                continue;
            }

            double lineWidth = MeasureRichLineWidth(line, pageOptions);
            double textX = ResolveFreeTextAppearanceLineX(annotation.TextAlign, annotation.Padding, availableWidth, lineWidth);
            AppendFreeTextAppearanceLine(
                sb,
                line,
                textX,
                Math.Max(0D, baseline),
                resolvedTextColor,
                pageOptions,
                ensureFont,
                resources);
            baseline -= effectiveLineHeight;
        }

        sb.Append("Q\n");
        fontResources = resources;
        return sb.ToString();
    }

    private static bool CanEncodeFreeTextAppearanceLine(IReadOnlyList<RichSeg> line, PdfOptions pageOptions) {
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (segment.LeadingSpace && !CanEncodeAppearanceText(" ", segment.Font, pageOptions)) {
                return false;
            }

            if (!CanEncodeAppearanceText(segment.Text, segment.Font, pageOptions)) {
                return false;
            }
        }

        return true;
    }

    private static bool CanEncodeAppearanceText(string text, PdfStandardFont font, PdfOptions pageOptions) {
        if (text.Length == 0) {
            return true;
        }

        if (pageOptions.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return CanWriteWithEmbeddedFont(text, fontProgram);
        }

        if (pageOptions.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return CanWriteWithEmbeddedFont(text, cffFontProgram);
        }

        return PdfWinAnsiEncoding.CanEncode(text, out _);
    }

    private static void AppendFreeTextAppearanceLine(
        StringBuilder sb,
        IReadOnlyList<RichSeg> line,
        double x,
        double y,
        PdfColor defaultTextColor,
        PdfOptions pageOptions,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        List<(string Name, int Id)> fontResources) {
        double xCursor = 0D;
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            string fontResource = GetFreeTextAppearanceFontResourceName(segment.Font, pageOptions);
            RegisterAppearanceFontResource(fontResources, fontResource, ensureFont(segment.Font, pageOptions));
            double runFontSize = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
            double textRise = TextRiseForBaseline(segment.FontSize, segment.Baseline);
            if (segment.LeadingSpace) {
                double gap = segment.LeadingAdvance > 0D
                    ? segment.LeadingAdvance
                    : MeasureRichText(" ", segment.Font, segment.FontSize, segment.Baseline, pageOptions);
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
                    pageOptions);
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
                    pageOptions);
                xCursor += MeasureRichText(segment.Text, segment.Font, segment.FontSize, segment.Baseline, pageOptions);
            }
        }
    }

    private static void AppendFreeTextAppearanceTextSegment(
        StringBuilder sb,
        string fontResource,
        double fontSize,
        double textRise,
        double x,
        double y,
        string text,
        PdfStandardFont font,
        PdfColor color,
        PdfOptions pageOptions) {
        sb.Append("BT /")
            .Append(PdfSyntaxEscaper.Name(fontResource))
            .Append(' ')
            .Append(FormatAppearanceNumber(fontSize))
            .Append(" Tf ");
        if (Math.Abs(textRise) > 0.0001D) {
            sb.Append(FormatAppearanceNumber(textRise))
                .Append(" Ts ");
        }

        sb.Append(FormatAppearanceColor(color))
            .Append(" rg ")
            .Append(FormatAppearanceNumber(x))
            .Append(' ')
            .Append(FormatAppearanceNumber(y))
            .Append(" Td <")
            .Append(EncodeTextHex(text, font, pageOptions))
            .Append("> Tj ET\n");
    }

    private static string GetFreeTextAppearanceFontResourceName(PdfStandardFont font, PdfOptions pageOptions) =>
        font == PdfStandardFont.Helvetica
            ? "Helv"
            : GetStandardFontResourceName(font, ChooseNormal(pageOptions.DefaultFont));

    private static void RegisterAppearanceFontResource(List<(string Name, int Id)> fontResources, string name, int id) {
        for (int index = 0; index < fontResources.Count; index++) {
            if (string.Equals(fontResources[index].Name, name, StringComparison.Ordinal)) {
                return;
            }
        }

        fontResources.Add((name, id));
    }

    private static double ResolveFreeTextAppearanceLineX(PdfAlign textAlign, double padding, double availableWidth, double lineWidth) {
        return textAlign switch {
            PdfAlign.Center => padding + Math.Max(0D, (availableWidth - lineWidth) / 2D),
            PdfAlign.Right => padding + Math.Max(0D, availableWidth - lineWidth),
            _ => padding
        };
    }

    private static string FormatAppearanceColor(PdfColor color) =>
        FormatAppearanceNumber(color.R) + " " + FormatAppearanceNumber(color.G) + " " + FormatAppearanceNumber(color.B);

    private static string FormatAppearanceNumber(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);
}
