namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double[] BuildPageTextLineBaselines(
        System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<TextRun>> lines,
        double firstBaseline,
        double defaultFontSize) {
        var baselines = new double[lines.Count];
        double baseline = firstBaseline;
        for (int index = 0; index < lines.Count; index++) {
            baselines[index] = baseline;
            baseline -= GetPageTextLineLeading(lines[index], defaultFontSize);
        }

        return baselines;
    }

    private static double GetPageTextLineLeading(System.Collections.Generic.IReadOnlyList<TextRun> line, double defaultFontSize) {
        double leading = defaultFontSize * 1.2D;
        foreach (TextRun run in line) {
            double requestedFontSize = run.FontSize ?? defaultFontSize;
            leading = Math.Max(leading, EffectiveRichFontSize(requestedFontSize, run.Baseline) * 1.2D);
        }

        return leading;
    }

    private static void AppendPageTextRunDecorations(
        StringBuilder sb,
        System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<TextRun>> lines,
        double[] baselines,
        PdfStandardFont baseFont,
        double defaultFontSize,
        PdfColor? defaultColor,
        double x,
        PdfOptions options,
        double? lineBoxWidth,
        PdfAlign align) {
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            System.Collections.Generic.IReadOnlyList<TextRun> line = lines[lineIndex];
            double lineWidth = MeasurePageTextLineRuns(line, baseFont, defaultFontSize, options);
            double dx = lineBoxWidth.HasValue
                ? align == PdfAlign.Center
                    ? Math.Max(0D, (lineBoxWidth.Value - lineWidth) / 2D)
                    : align == PdfAlign.Right
                        ? Math.Max(0D, lineBoxWidth.Value - lineWidth)
                        : 0D
                : 0D;
            double cursorX = x + dx;

            foreach (TextRun run in line) {
                string text = run.Text ?? string.Empty;
                if (text.Length == 0) {
                    continue;
                }

                PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
                PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                    ? resolvedNamedFont
                    : null;
                double requestedFontSize = run.FontSize ?? defaultFontSize;
                double effectiveFontSize = EffectiveRichFontSize(requestedFontSize, run.Baseline);
                double textRise = TextRiseForBaseline(requestedFontSize, run.Baseline);
                double width = MeasureRichText(text, runFont, namedFont, requestedFontSize, run.Baseline, options);
                PdfColor runColor = ResolvePageTextColor(run.Color ?? defaultColor, options);

                if (run.BackgroundColor.HasValue && width > 0D) {
                    double ascender = GetAscenderForOptions(runFont, namedFont, effectiveFontSize, options);
                    double descender = GetDescenderForOptions(runFont, namedFont, effectiveFontSize, options);
                    double paddingY = Math.Max(0.35D, effectiveFontSize * 0.04D);
                    new ContentStreamBuilder(sb)
                        .SaveState()
                        .FillColor(run.BackgroundColor.Value)
                        .Rectangle(cursorX, baselines[lineIndex] + textRise - descender - paddingY, width, ascender + descender + (paddingY * 2D))
                        .FillPath()
                        .RestoreState();
                }

                if (width > 0D && (run.Underline || run.Strike)) {
                    double decorationWidth = Math.Max(0.45D, effectiveFontSize * 0.055D);
                    if (run.Underline) {
                        AppendPageTextDecorationLine(sb, cursorX, cursorX + width, baselines[lineIndex] + textRise - Math.Max(0.8D, effectiveFontSize * 0.1D), decorationWidth, runColor);
                    }
                    if (run.Strike) {
                        AppendPageTextDecorationLine(sb, cursorX, cursorX + width, baselines[lineIndex] + textRise + (effectiveFontSize * 0.28D), decorationWidth, runColor);
                    }
                }

                cursorX += width;
            }
        }
    }

    private static void AppendPageTextDecorationLine(StringBuilder sb, double x1, double x2, double y, double width, PdfColor color) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(width)
            .MoveTo(x1, y)
            .LineTo(x2, y)
            .StrokePath()
            .RestoreState();
    }
}
