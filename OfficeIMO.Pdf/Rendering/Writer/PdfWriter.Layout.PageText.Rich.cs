namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void GetPageTextVerticalBoundsOrBaseline(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs,
        PdfStandardFont baseFont,
        double defaultFontSize,
        PdfOptions options,
        double firstBaseline,
        out double bottom,
        out double top) {
        if (!TryGetPageTextVerticalBounds(runs, baseFont, defaultFontSize, options, firstBaseline, out bottom, out top)) {
            bottom = firstBaseline;
            top = firstBaseline;
        }
    }

    private static bool TryGetPageTextVerticalBounds(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs,
        PdfStandardFont baseFont,
        double defaultFontSize,
        PdfOptions options,
        double firstBaseline,
        out double bottom,
        out double top) {
        var lines = BuildPageTextLineRuns(runs);
        double[] baselines = BuildPageTextLineBaselines(lines, firstBaseline, defaultFontSize);
        bottom = double.PositiveInfinity;
        top = double.NegativeInfinity;

        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            foreach (PdfTextRun run in lines[lineIndex]) {
                if (string.IsNullOrEmpty(run.Text)) {
                    continue;
                }

                PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
                PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                    ? resolvedNamedFont
                    : null;
                double requestedFontSize = run.FontSize ?? defaultFontSize;
                double effectiveFontSize = EffectiveRichFontSize(requestedFontSize, run.Baseline);
                double textRise = TextRiseForBaseline(requestedFontSize, run.Baseline);
                double runBaseline = baselines[lineIndex] + textRise;
                double ascender = GetAscenderForOptions(runFont, namedFont, effectiveFontSize, options);
                double descender = GetDescenderForOptions(runFont, namedFont, effectiveFontSize, options);
                double runBottom = runBaseline - descender;
                double runTop = runBaseline + ascender;
                double width = MeasureRichText(run.Text, runFont, namedFont, requestedFontSize, run.Baseline, options);
                if (width > 0D && run.BackgroundColor.HasValue) {
                    double paddingY = System.Math.Max(0.35D, effectiveFontSize * 0.04D);
                    runBottom -= paddingY;
                    runTop += paddingY;
                }

                if (width > 0D && (run.Underline || run.Strike)) {
                    double decorationWidth = System.Math.Max(0.45D, effectiveFontSize * 0.055D);
                    if (run.Underline) {
                        double underlineY = runBaseline - System.Math.Max(0.8D, effectiveFontSize * 0.1D);
                        IncludeDecorationVerticalBounds(run.UnderlineStyle, underlineY, decorationWidth, ref runBottom, ref runTop);
                    }

                    if (run.Strike) {
                        double strikeY = runBaseline + (effectiveFontSize * 0.28D);
                        IncludeDecorationVerticalBounds(run.StrikeStyle, strikeY, decorationWidth, ref runBottom, ref runTop);
                    }
                }

                bottom = System.Math.Min(bottom, runBottom);
                top = System.Math.Max(top, runTop);
            }
        }

        return !double.IsPositiveInfinity(bottom) && !double.IsNegativeInfinity(top);
    }

    private static void IncludeDecorationVerticalBounds(OfficeIMO.Drawing.OfficeTextDecorationStyle style, double y, double width, ref double bottom, ref double top) {
        double extra = width / 2D;
        if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Double) {
            extra += Math.Max(width * 1.8D, 0.8D) / 2D;
        } else if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            extra += Math.Max(width * 1.5D, 0.75D);
        }

        bottom = Math.Min(bottom, y - extra);
        top = Math.Max(top, y + extra);
    }

    private static double[] BuildPageTextLineBaselines(
        System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTextRun>> lines,
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

    private static double GetPageTextLineLeading(System.Collections.Generic.IReadOnlyList<PdfTextRun> line, double defaultFontSize) {
        double leading = defaultFontSize * 1.2D;
        foreach (PdfTextRun run in line) {
            double requestedFontSize = run.FontSize ?? defaultFontSize;
            leading = Math.Max(leading, EffectiveRichFontSize(requestedFontSize, run.Baseline) * 1.2D);
        }

        return leading;
    }

    private static void AppendPageTextRunDecorations(
        StringBuilder sb,
        System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTextRun>> lines,
        double[] baselines,
        PdfStandardFont baseFont,
        double defaultFontSize,
        PdfColor? defaultColor,
        double x,
        PdfOptions options,
        double? lineBoxWidth,
        PdfAlign align) {
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            System.Collections.Generic.IReadOnlyList<PdfTextRun> line = lines[lineIndex];
            double lineWidth = MeasurePageTextLineRuns(line, baseFont, defaultFontSize, options);
            double dx = lineBoxWidth.HasValue
                ? align == PdfAlign.Center
                    ? Math.Max(0D, (lineBoxWidth.Value - lineWidth) / 2D)
                    : align == PdfAlign.Right
                        ? Math.Max(0D, lineBoxWidth.Value - lineWidth)
                        : 0D
                : 0D;
            double cursorX = x + dx;

            foreach (PdfTextRun run in line) {
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
                PdfColor decorationColor = ResolvePageTextColor(run.DecorationColor ?? run.Color ?? defaultColor, options);

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
                        AppendPageTextDecorationLine(sb, cursorX, cursorX + width, baselines[lineIndex] + textRise - Math.Max(0.8D, effectiveFontSize * 0.1D), decorationWidth, decorationColor, run.UnderlineStyle);
                    }
                    if (run.Strike) {
                        AppendPageTextDecorationLine(sb, cursorX, cursorX + width, baselines[lineIndex] + textRise + (effectiveFontSize * 0.28D), decorationWidth, decorationColor, run.StrikeStyle);
                    }
                }

                cursorX += width;
            }
        }
    }

    private static void AppendPageTextDecorationLine(StringBuilder sb, double x1, double x2, double y, double width, PdfColor color, OfficeIMO.Drawing.OfficeTextDecorationStyle style) {
        if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.None || x2 <= x1) {
            return;
        }

        if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Double) {
            double separation = Math.Max(width * 1.8D, 0.8D);
            AppendPageTextDecorationLine(sb, x1, x2, y - (separation / 2D), width, color, OfficeIMO.Drawing.OfficeTextDecorationStyle.Single);
            AppendPageTextDecorationLine(sb, x1, x2, y + (separation / 2D), width, color, OfficeIMO.Drawing.OfficeTextDecorationStyle.Single);
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(width);
        if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Dashed) {
            content.StrokeDash(Math.Max(width * 5D, 2D), Math.Max(width * 3D, 1D));
        } else if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Dotted) {
            content.LineCap(1).StrokeDash(Math.Max(width, 0.5D), Math.Max(width * 3D, 1D));
        }

        if (style == OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            double step = Math.Max(width * 5D, 2D);
            double amplitude = Math.Max(width * 1.5D, 0.75D);
            content.MoveTo(x1, y);
            double cursor = x1;
            bool raised = true;
            while (cursor < x2) {
                double next = Math.Min(x2, cursor + step);
                content.LineTo(next, y + (raised ? amplitude : -amplitude));
                cursor = next;
                raised = !raised;
            }
        } else {
            content.MoveTo(x1, y).LineTo(x2, y);
        }

        content.StrokePath().RestoreState();
    }
}
