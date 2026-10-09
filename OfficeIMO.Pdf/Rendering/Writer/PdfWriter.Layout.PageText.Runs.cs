using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> BuildPageTextRuns(string text, PdfStandardFont font, double fontSize, PdfColor? color, PdfOptions opts, string? fontFamily = null) {
        (bool bold, bool italic) = GetPageTextFontStyle(font);
        var run = new PdfTextRun(
            text,
            bold: bold,
            underline: false,
            color: color,
            italic: italic,
            strike: false,
            fontSize: fontSize,
            font: ChooseNormal(font),
            fontFamily: fontFamily);
        return NormalizeFallbackRuns(new[] { run }, ChooseNormal(font), opts);
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> BuildPageTextRunsFromSegments(
        System.Collections.Generic.IReadOnlyList<FooterSegment> segments,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        PdfColor? color,
        PdfOptions opts,
        string? fontFamily) {
        (bool bold, bool italic) = GetPageTextFontStyle(font);
        var runs = new System.Collections.Generic.List<PdfTextRun>(segments.Count);
        foreach (FooterSegment segment in segments) {
            string text = segment.Kind switch {
                FooterSegmentKind.Text => segment.Text ?? string.Empty,
                FooterSegmentKind.PageNumber => FormatPageNumber(page, opts.PageNumberStyle),
                FooterSegmentKind.TotalPages => FormatPageNumber(pages, opts.PageNumberStyle),
                FooterSegmentKind.DocumentPages => FormatPageNumber(documentPages, opts.PageNumberStyle),
                _ => throw new System.ArgumentOutOfRangeException(nameof(segments), segment.Kind, "PDF header/footer segment kind is not supported.")
            };

            if (segment.StyledRun != null) {
                runs.Add(CreateStyledTextRun(text, segment.StyledRun, segment.StyledRun.Font, fontFamily));
            } else {
                runs.Add(new PdfTextRun(
                    text,
                    bold: bold,
                    underline: false,
                    color: color,
                    italic: italic,
                    strike: false,
                    fontSize: fontSize,
                    font: ChooseNormal(font),
                    fontFamily: fontFamily));
            }
        }

        return NormalizeFallbackRuns(runs, ChooseNormal(font), opts);
    }

    private static bool TryResolvePageTextNamedFont(PdfOptions options, string? fontFamily, PdfStandardFont font, out PdfNamedFontFace namedFont) {
        (bool bold, bool italic) = GetPageTextFontStyle(font);
        return options.TryResolveNamedFontFace(fontFamily, bold, italic, out namedFont);
    }

    private static (bool Bold, bool Italic) GetPageTextFontStyle(PdfStandardFont font) {
        bool bold = font == PdfStandardFont.HelveticaBold ||
            font == PdfStandardFont.HelveticaBoldOblique ||
            font == PdfStandardFont.TimesBold ||
            font == PdfStandardFont.TimesBoldItalic ||
            font == PdfStandardFont.CourierBold ||
            font == PdfStandardFont.CourierBoldOblique;
        bool italic = font == PdfStandardFont.HelveticaOblique ||
            font == PdfStandardFont.HelveticaBoldOblique ||
            font == PdfStandardFont.TimesItalic ||
            font == PdfStandardFont.TimesBoldItalic ||
            font == PdfStandardFont.CourierOblique ||
            font == PdfStandardFont.CourierBoldOblique;
        return (bold, italic);
    }

    private static double MeasurePageTextRuns(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions opts) {
        double width = 0D;
        foreach (System.Collections.Generic.IReadOnlyList<PdfTextRun> line in BuildPageTextLineRuns(runs)) {
            width = Math.Max(width, MeasurePageTextLineRuns(line, baseFont, fontSize, opts));
        }

        return width;
    }

    private static double MeasurePageTextLineRuns(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions opts) {
        double width = 0D;
        foreach (PdfTextRun run in runs) {
            width += run.HorizontalOffset;
            PdfNamedFontFace? namedFont = opts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            width += run.InlineElement?.Width ?? MeasureRichText(run.Text ?? string.Empty, ResolvePageTextRunFont(run, baseFont), namedFont, run.FontSize ?? fontSize, run.Baseline, opts, run.FeatureSettings, run.HorizontalTextScaling, run.CharacterSpacing);
        }

        return width;
    }

    private static System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTextRun>> BuildPageTextLineRuns(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTextRun>>();
        var current = new System.Collections.Generic.List<PdfTextRun>();
        lines.Add(current);

        foreach (PdfTextRun run in runs) {
            if (run.InlineElement != null) {
                current.Add(run);
                continue;
            }

            string text = run.Text ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char ch = text[index];
                if (ch != '\r' && ch != '\n') {
                    continue;
                }

                if (index > segmentStart) {
                    current.Add(CreateStyledTextRun(text.Substring(segmentStart, index - segmentStart), run, run.Font));
                }

                if (lines.Count >= MaximumTextLayoutLines)
                    throw new System.IO.InvalidDataException("PDF page text layout exceeds the 100,000-line limit.");
                current = new System.Collections.Generic.List<PdfTextRun>();
                lines.Add(current);
                if (ch == '\r' && index + 1 < text.Length && text[index + 1] == '\n') {
                    index++;
                }

                segmentStart = index + 1;
            }

            if (segmentStart < text.Length) {
                current.Add(CreateStyledTextRun(text.Substring(segmentStart), run, run.Font));
            }
        }

        return lines;
    }

    private static void AppendPageTextRuns(
        StringBuilder sb,
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs,
        PdfStandardFont baseFont,
        string baseFontResource,
        System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources,
        System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources,
        double fontSize,
        PdfColor? color,
        double x,
        double y,
        PdfOptions opts,
        double? lineBoxWidth = null,
        PdfAlign align = PdfAlign.Left) {
        var lines = BuildPageTextLineRuns(runs);
        double[] baselines = BuildPageTextLineBaselines(lines, y, fontSize);
        AppendPageTextRunDecorations(sb, lines, baselines, baseFont, fontSize, color, x, opts, lineBoxWidth, align);

        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(baseFontResource, fontSize, opts.NeedsSyntheticOblique(baseFont))
            .FillColor(ResolvePageTextColor(color, opts))
            .TextLeading(fontSize * 1.2D);

        double currentTextRise = 0D;
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            System.Collections.Generic.IReadOnlyList<PdfTextRun> line = lines[lineIndex];
            double dx = 0D;
            if (lineBoxWidth.HasValue) {
                double lineWidth = MeasurePageTextLineRuns(line, baseFont, fontSize, opts);
                if (align == PdfAlign.Center) {
                    dx = Math.Max(0D, (lineBoxWidth.Value - lineWidth) / 2D);
                } else if (align == PdfAlign.Right) {
                    dx = Math.Max(0D, lineBoxWidth.Value - lineWidth);
                }
            }
            if (lineIndex > 0 && Math.Abs(currentTextRise) > 0.0001D) {
                content.TextRise(0D);
                currentTextRise = 0D;
            }
            double cursorX = x + dx;
            foreach (PdfTextRun run in line) {
                string text = run.Text ?? string.Empty;
                if (text.Length == 0) {
                    continue;
                }

                cursorX += run.HorizontalOffset;

                PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
                PdfNamedFontFace? namedFont = opts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                    ? resolvedNamedFont
                    : null;
                string fontResource = ResolvePageTextFontResource(fontResources, namedFontResources, runFont, namedFont);
                double requestedFontSize = run.FontSize ?? fontSize;
                double runFontSize = EffectiveRichFontSize(requestedFontSize, run.Baseline);
                double textRise = TextRiseForBaseline(requestedFontSize, run.Baseline);
                content.Font(fontResource, runFontSize, opts.NeedsSyntheticOblique(runFont, namedFont));
                if (Math.Abs(textRise - currentTextRise) > 0.0001D) {
                    content.TextRise(textRise);
                    currentTextRise = textRise;
                }
                ApplyRichTextSpacing(content, run.HorizontalTextScaling, run.CharacterSpacing);
                content
                    .TextMatrix(cursorX, baselines[lineIndex])
                    .FillColor(ResolvePageTextColor(run.Color ?? color, opts))
                    .ShowText(EncodeTextShowCommand(text, runFont, namedFont, opts, run.FeatureSettings, run.TextDirection), runFontSize);
                ResetRichTextSpacing(content, run.HorizontalTextScaling, run.CharacterSpacing);
                cursorX += MeasureRichText(text, runFont, namedFont, requestedFontSize, run.Baseline, opts, run.FeatureSettings, run.HorizontalTextScaling, run.CharacterSpacing);
            }
        }

        if (Math.Abs(currentTextRise) > 0.0001D) {
            content.TextRise(0D);
        }

        content.EndText();
    }

    private static PdfStandardFont ResolvePageTextRunFont(PdfTextRun run, PdfStandardFont baseFont) {
        PdfStandardFont runBaseFont = run.Font ?? baseFont;
        return (run.Bold && run.Italic)
            ? ChooseBoldItalic(ChooseNormal(runBaseFont))
            : run.Bold
                ? ChooseBold(ChooseNormal(runBaseFont))
                : run.Italic
                    ? ChooseItalic(ChooseNormal(runBaseFont))
                    : runBaseFont;
    }

    private static string ResolvePageTextFontResource(System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, PdfStandardFont font) {
        if (!fontResources.TryGetValue(font, out string? fontResource)) {
            throw new InvalidOperationException("PDF page text font resource was not registered before rendering.");
        }

        return fontResource;
    }

    private static string ResolvePageTextFontResource(
        System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources,
        System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources,
        PdfStandardFont font,
        PdfNamedFontFace? namedFont) {
        if (namedFont.HasValue && namedFontResources.TryGetValue(namedFont.Value, out string? namedFontResource)) {
            return namedFontResource;
        }

        return ResolvePageTextFontResource(fontResources, font);
    }

    private static void AppendPageText(StringBuilder sb, string text, PdfStandardFont font, string fontResource, double fontSize, PdfColor? color, double x, double y, PdfOptions opts) {
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(fontResource, fontSize, opts.NeedsSyntheticOblique(font))
            .FillColor(ResolvePageTextColor(color, opts));

        content
            .TextMatrix(x, y)
            .ShowText(EncodeTextShowCommand(text, font, opts), fontSize)
            .EndText();
    }

    private static PdfColor ResolvePageTextColor(PdfColor? color, PdfOptions opts) =>
        color ?? opts.DefaultTextColor ?? PdfColor.Black;

    private static string FormatPageText(string format, int page, int pages, int documentPages, PdfPageNumberStyle style) {
        string pageText = FormatPageNumber(page, style);
        string pagesText = FormatPageNumber(pages, style);
        string documentPagesText = FormatPageNumber(documentPages, style);
        return format
            .Replace("{page}", pageText)
            .Replace("{pages}", pagesText)
            .Replace("{documentpages}", documentPagesText);
    }

    private static string FormatPageNumber(int number, PdfPageNumberStyle style) {
        return PdfPageNumberFormatter.Format(number, style);
    }

}
