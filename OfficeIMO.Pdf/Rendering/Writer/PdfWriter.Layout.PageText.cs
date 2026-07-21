using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFooter(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont footerFont, string footerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources) {
        System.Collections.Generic.IReadOnlyList<TextRun> runs;
        var footerSegments = opts.GetFooterSegmentsForPage(variantPage);
        var footerZones = opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(footerZones)) {
            return BuildPageTextZones(opts, footerZones, variantPage, page, pages, documentPages, footerFont, fontResources, namedFontResources, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (footerSegments != null && footerSegments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(footerSegments, page, pages, footerFont, opts.FooterFontSize, opts.FooterTextColor, opts, opts.FooterFontFamily);
        } else {
            string text = FormatPageText(opts.GetFooterFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
            runs = BuildPageTextRuns(text, footerFont, opts.FooterFontSize, opts.FooterTextColor, opts, opts.FooterFontFamily);
        }
        double textWidth = MeasurePageTextRuns(runs, footerFont, opts.FooterFontSize, opts);
        double imagesWidth = MeasureHeaderFooterImagesWidth(opts.GetFooterImagesForPage(variantPage), opts.FooterAlign);
        double shapesWidth = MeasureHeaderFooterShapesWidth(opts.GetFooterShapesForPage(variantPage), opts.FooterAlign);
        double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
        double x = AlignHeaderFooterGroup(opts, groupWidth, opts.FooterAlign);
        double y = opts.MarginBottom - opts.FooterOffsetY;
        PdfColor? footerColor = opts.FooterTextColor;
        var sb = new StringBuilder();
        AppendPageTextRuns(sb, runs, footerFont, footerFontResource, fontResources, namedFontResources, opts.FooterFontSize, footerColor, x, y, opts, textWidth, opts.FooterAlign);
        return sb.ToString();
    }

    private static string BuildHeader(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont headerFont, string headerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources) {
        System.Collections.Generic.IReadOnlyList<TextRun> runs;
        var headerSegments = opts.GetHeaderSegmentsForPage(variantPage);
        var headerZones = opts.GetHeaderZonesForPage(variantPage);
        if (HasPageTextZones(headerZones)) {
            return BuildPageTextZones(opts, headerZones, variantPage, page, pages, documentPages, headerFont, fontResources, namedFontResources, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (headerSegments != null && headerSegments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(headerSegments, page, pages, headerFont, opts.HeaderFontSize, opts.HeaderTextColor, opts, opts.HeaderFontFamily);
        } else {
            string text = FormatPageText(opts.GetHeaderFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
            runs = BuildPageTextRuns(text, headerFont, opts.HeaderFontSize, opts.HeaderTextColor, opts, opts.HeaderFontFamily);
        }

        double textWidth = MeasurePageTextRuns(runs, headerFont, opts.HeaderFontSize, opts);
        double imagesWidth = MeasureHeaderFooterImagesWidth(opts.GetHeaderImagesForPage(variantPage), opts.HeaderAlign);
        double shapesWidth = MeasureHeaderFooterShapesWidth(opts.GetHeaderShapesForPage(variantPage), opts.HeaderAlign);
        double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
        double x = AlignHeaderFooterGroup(opts, groupWidth, opts.HeaderAlign);
        double y = opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY;
        PdfColor? headerColor = opts.HeaderTextColor;

        var sb = new StringBuilder();
        AppendPageTextRuns(sb, runs, headerFont, headerFontResource, fontResources, namedFontResources, opts.HeaderFontSize, headerColor, x, y, opts, textWidth, opts.HeaderAlign);
        return sb.ToString();
    }

    private static void EnsurePageTextFontResources(
        PdfOptions opts,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        bool isHeader,
        Func<PdfStandardFont, string, string> ensureFontResource,
        Action<PdfNamedFontFace> ensureNamedFontResource) {
        string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;
        var zones = isHeader ? opts.GetHeaderZonesForPage(variantPage) : opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(zones)) {
            EnsurePageTextZoneFontResources(opts, zones, page, pages, documentPages, font, fontSize, fontFamily, ensureFontResource, ensureNamedFontResource);
            return;
        }

        System.Collections.Generic.IReadOnlyList<TextRun> runs;
        var segments = isHeader ? opts.GetHeaderSegmentsForPage(variantPage) : opts.GetFooterSegmentsForPage(variantPage);
        if (segments != null && segments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(segments, page, pages, font, fontSize, color: null, opts, fontFamily);
        } else {
            string text = FormatPageText(isHeader ? opts.GetHeaderFormatForPage(variantPage) : opts.GetFooterFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
            runs = BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily);
        }

        EnsurePageTextRunFontResources(runs, font, opts, ensureFontResource, ensureNamedFontResource);
    }

    private static void EnsurePageTextZoneFontResources(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        string? fontFamily,
        Func<PdfStandardFont, string, string> ensureFontResource,
        Action<PdfNamedFontFace> ensureNamedFontResource) {
        if (!string.IsNullOrEmpty(zones.Left)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Left!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts, fontFamily), font, opts, ensureFontResource, ensureNamedFontResource);
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts, fontFamily), font, opts, ensureFontResource, ensureNamedFontResource);
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts, fontFamily), font, opts, ensureFontResource, ensureNamedFontResource);
        }
    }

    private static void EnsurePageTextRunFontResources(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, PdfOptions opts, Func<PdfStandardFont, string, string> ensureFontResource, Action<PdfNamedFontFace> ensureNamedFontResource) {
        PdfStandardFont normalFont = ChooseNormal(opts.DefaultFont);
        foreach (TextRun run in runs) {
            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            if (opts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace namedFont)) {
                ensureNamedFontResource(namedFont);
            } else {
                ensureFontResource(runFont, GetStandardFontResourceName(runFont, normalFont));
            }
        }
    }

    private static bool HasPageTextZones((string? Left, string? Center, string? Right) zones) =>
        !string.IsNullOrEmpty(zones.Left) ||
        !string.IsNullOrEmpty(zones.Center) ||
        !string.IsNullOrEmpty(zones.Right);

    private static string BuildPageTextZones(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources,
        System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources,
        double fontSize,
        PdfColor? color,
        double offset,
        bool isHeader) {
        double y = isHeader ? opts.PageHeight - opts.MarginTop + offset : opts.MarginBottom - offset;
        var sb = new StringBuilder();
        var zoneLayouts = BuildPageTextZoneLayouts(opts, zones, variantPage, page, pages, documentPages, font, fontSize, isHeader);
        foreach (var zone in zoneLayouts) {
            string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;
            System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildPageTextRuns(zone.Text, font, fontSize, color, opts, fontFamily);
            PdfNamedFontFace? namedFont = TryResolvePageTextNamedFont(opts, fontFamily, font, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            string baseFontResource = ResolvePageTextFontResource(fontResources, namedFontResources, font, namedFont);
            AppendPageTextRuns(sb, runs, font, baseFontResource, fontResources, namedFontResources, fontSize, color, zone.X, y, opts, zone.TextWidth, zone.Align);
        }

        return sb.ToString();
    }

    private static System.Collections.Generic.List<PageTextZoneLayout> BuildPageTextZoneLayouts(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        bool isHeader) {
        double contentLeft = opts.MarginLeft;
        double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        var layouts = new System.Collections.Generic.List<PageTextZoneLayout>();
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images = isHeader
            ? opts.GetHeaderImagesForPage(variantPage)
            : opts.GetFooterImagesForPage(variantPage);
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes = isHeader
            ? opts.GetHeaderShapesForPage(variantPage)
            : opts.GetFooterShapesForPage(variantPage);
        string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;

        if (!string.IsNullOrEmpty(zones.Left)) {
            string text = FormatPageText(zones.Left!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily), font, fontSize, opts);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Left);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Left);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            layouts.Add(new PageTextZoneLayout(text, contentLeft, textWidth, PdfAlign.Left, contentLeft, occupiedWidth));
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            string text = FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily), font, fontSize, opts);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Center);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Center);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            double occupiedX = contentLeft + ((contentWidth - occupiedWidth) / 2);
            layouts.Add(new PageTextZoneLayout(text, occupiedX, textWidth, PdfAlign.Center, occupiedX, occupiedWidth));
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            string text = FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily), font, fontSize, opts);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Right);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Right);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            double occupiedX = contentLeft + contentWidth - occupiedWidth;
            layouts.Add(new PageTextZoneLayout(text, occupiedX, textWidth, PdfAlign.Right, occupiedX, occupiedWidth));
        }

        ValidatePageTextZoneLayouts(layouts, contentLeft, contentLeft + contentWidth, isHeader);
        return layouts;
    }

    private static void ValidatePageTextZoneLayouts(System.Collections.Generic.List<PageTextZoneLayout> layouts, double contentLeft, double contentRight, bool isHeader) {
        const double tolerance = 0.01D;
        const double minimumGap = 2D;
        string scope = isHeader ? "header" : "footer";
        foreach (var zone in layouts) {
            if (zone.OccupiedX < contentLeft - tolerance || zone.OccupiedX + zone.OccupiedWidth > contentRight + tolerance) {
                throw new ArgumentException("PDF " + scope + " zone content must fit inside the page content width.");
            }
        }

        var ordered = layouts.OrderBy(zone => zone.OccupiedX).ToList();
        for (int i = 1; i < ordered.Count; i++) {
            var previous = ordered[i - 1];
            var current = ordered[i];
            if (previous.OccupiedX + previous.OccupiedWidth + minimumGap > current.OccupiedX + tolerance) {
                throw new ArgumentException("PDF " + scope + " zones must not overlap.");
            }
        }
    }

    private static double MeasureHeaderFooterTextWidth(
        PdfOptions opts,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfAlign align,
        bool isHeader) {
        PdfStandardFont font = isHeader ? opts.HeaderFont : opts.FooterFont;
        double fontSize = isHeader ? opts.HeaderFontSize : opts.FooterFontSize;
        var zones = isHeader ? opts.GetHeaderZonesForPage(variantPage) : opts.GetFooterZonesForPage(variantPage);
        string? text = align switch {
            PdfAlign.Center => zones.Center,
            PdfAlign.Right => zones.Right,
            _ => zones.Left
        };
        System.Collections.Generic.IReadOnlyList<TextRun>? runs = null;

        if (!string.IsNullOrEmpty(text)) {
            text = FormatPageText(text!, page, pages, documentPages, opts.PageNumberStyle);
        } else if ((isHeader ? opts.HasHeaderTextContentForPage(variantPage) : opts.HasFooterTextContentForPage(variantPage)) &&
                   !HasPageTextZones(zones) &&
                   align == (isHeader ? opts.HeaderAlign : opts.FooterAlign)) {
            var segments = isHeader ? opts.GetHeaderSegmentsForPage(variantPage) : opts.GetFooterSegmentsForPage(variantPage);
            if (segments != null && segments.Count > 0) {
                runs = BuildPageTextRunsFromSegments(
                    segments,
                    page,
                    pages,
                    font,
                    fontSize,
                    color: null,
                    opts,
                    isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily);
            } else {
                text = FormatPageText(
                    isHeader ? opts.GetHeaderFormatForPage(variantPage) : opts.GetFooterFormatForPage(variantPage),
                    page,
                    pages,
                    documentPages,
                    opts.PageNumberStyle);
            }
        }

        if (runs == null && !string.IsNullOrEmpty(text)) {
            runs = BuildPageTextRuns(text!, font, fontSize, color: null, opts, isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily);
        }

        return runs == null ? 0D : MeasurePageTextRuns(runs, font, fontSize, opts);
    }

    private readonly record struct PageTextZoneLayout(
        string Text,
        double X,
        double TextWidth,
        PdfAlign Align,
        double OccupiedX,
        double OccupiedWidth);

    private static System.Collections.Generic.IReadOnlyList<TextRun> BuildPageTextRuns(string text, PdfStandardFont font, double fontSize, PdfColor? color, PdfOptions opts, string? fontFamily = null) {
        (bool bold, bool italic) = GetPageTextFontStyle(font);
        var run = new TextRun(
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

    private static System.Collections.Generic.IReadOnlyList<TextRun> BuildPageTextRunsFromSegments(
        System.Collections.Generic.IReadOnlyList<FooterSegment> segments,
        int page,
        int pages,
        PdfStandardFont font,
        double fontSize,
        PdfColor? color,
        PdfOptions opts,
        string? fontFamily) {
        (bool bold, bool italic) = GetPageTextFontStyle(font);
        var runs = new System.Collections.Generic.List<TextRun>(segments.Count);
        foreach (FooterSegment segment in segments) {
            string text = segment.Kind switch {
                FooterSegmentKind.Text => segment.Text ?? string.Empty,
                FooterSegmentKind.PageNumber => FormatPageNumber(page, opts.PageNumberStyle),
                FooterSegmentKind.TotalPages => FormatPageNumber(pages, opts.PageNumberStyle),
                _ => throw new System.ArgumentOutOfRangeException(nameof(segments), segment.Kind, "PDF header/footer segment kind is not supported.")
            };

            if (segment.StyledRun != null) {
                runs.Add(CreateStyledTextRun(text, segment.StyledRun, segment.StyledRun.Font, fontFamily));
            } else {
                runs.Add(new TextRun(
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

    private static double MeasurePageTextRuns(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions opts) {
        double width = 0D;
        foreach (System.Collections.Generic.IReadOnlyList<TextRun> line in BuildPageTextLineRuns(runs)) {
            width = Math.Max(width, MeasurePageTextLineRuns(line, baseFont, fontSize, opts));
        }

        return width;
    }

    private static double MeasurePageTextLineRuns(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions opts) {
        double width = 0D;
        foreach (TextRun run in runs) {
            PdfNamedFontFace? namedFont = opts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            width += run.InlineElement?.Width ?? MeasureRichText(run.Text ?? string.Empty, ResolvePageTextRunFont(run, baseFont), namedFont, run.FontSize ?? fontSize, run.Baseline, opts);
        }

        return width;
    }

    private static System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<TextRun>> BuildPageTextLineRuns(System.Collections.Generic.IReadOnlyList<TextRun> runs) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<TextRun>>();
        var current = new System.Collections.Generic.List<TextRun>();
        lines.Add(current);

        foreach (TextRun run in runs) {
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

                current = new System.Collections.Generic.List<TextRun>();
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
        System.Collections.Generic.IReadOnlyList<TextRun> runs,
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
            .Font(baseFontResource, fontSize)
            .FillColor(ResolvePageTextColor(color, opts))
            .TextLeading(fontSize * 1.2D);

        double currentTextRise = 0D;
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            System.Collections.Generic.IReadOnlyList<TextRun> line = lines[lineIndex];
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
            content.TextMatrix(x + dx, baselines[lineIndex]);
            foreach (TextRun run in line) {
                string text = run.Text ?? string.Empty;
                if (text.Length == 0) {
                    continue;
                }

                PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
                PdfNamedFontFace? namedFont = opts.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolvedNamedFont)
                    ? resolvedNamedFont
                    : null;
                string fontResource = ResolvePageTextFontResource(fontResources, namedFontResources, runFont, namedFont);
                double requestedFontSize = run.FontSize ?? fontSize;
                double runFontSize = EffectiveRichFontSize(requestedFontSize, run.Baseline);
                double textRise = TextRiseForBaseline(requestedFontSize, run.Baseline);
                content.Font(fontResource, runFontSize);
                if (Math.Abs(textRise - currentTextRise) > 0.0001D) {
                    content.TextRise(textRise);
                    currentTextRise = textRise;
                }
                content
                    .FillColor(ResolvePageTextColor(run.Color ?? color, opts))
                    .ShowText(EncodeTextShowCommand(text, runFont, namedFont, opts), runFontSize);
            }
        }

        if (Math.Abs(currentTextRise) > 0.0001D) {
            content.TextRise(0D);
        }

        content.EndText();
    }

    private static PdfStandardFont ResolvePageTextRunFont(TextRun run, PdfStandardFont baseFont) {
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
            .Font(fontResource, fontSize)
            .FillColor(ResolvePageTextColor(color, opts));

        content
            .TextMatrix(x, y)
            .ShowText(EncodeTextShowCommand(text, font, opts), fontSize)
            .EndText();
    }

    private static PdfColor ResolvePageTextColor(PdfColor? color, PdfOptions opts) =>
        color ?? opts.DefaultTextColor ?? PdfColor.Black;

    private static string BuildPageTextFromSegments(System.Collections.Generic.IReadOnlyList<FooterSegment> segments, int page, int pages, PdfPageNumberStyle style) {
        var sb = new StringBuilder();
        foreach (var segment in segments) {
            switch (segment.Kind) {
                case FooterSegmentKind.Text:
                    sb.Append(segment.Text);
                    break;
                case FooterSegmentKind.PageNumber:
                    sb.Append(FormatPageNumber(page, style));
                    break;
                case FooterSegmentKind.TotalPages:
                    sb.Append(FormatPageNumber(pages, style));
                    break;
            }
        }

        return sb.ToString();
    }

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
        Guard.PageNumberStyle(style, nameof(style));
        if (number < 1) {
            throw new ArgumentOutOfRangeException(nameof(number), "PDF page number must be positive.");
        }

        switch (style) {
            case PdfPageNumberStyle.Arabic:
                return number.ToString(CultureInfo.InvariantCulture);
            case PdfPageNumberStyle.LowerRoman:
                return ToRoman(number).ToLowerInvariant();
            case PdfPageNumberStyle.UpperRoman:
                return ToRoman(number);
            case PdfPageNumberStyle.LowerLetter:
                return ToLetters(number, upper: false);
            case PdfPageNumberStyle.UpperLetter:
                return ToLetters(number, upper: true);
            default:
                throw new ArgumentException("PDF page number style must be Arabic, LowerRoman, UpperRoman, LowerLetter, or UpperLetter.", nameof(style));
        }
    }

    private static string ToRoman(int number) {
        var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var numerals = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        int remaining = number;
        for (int i = 0; i < values.Length; i++) {
            while (remaining >= values[i]) {
                sb.Append(numerals[i]);
                remaining -= values[i];
            }
        }

        return sb.ToString();
    }

    private static string ToLetters(int number, bool upper) {
        var chars = new System.Collections.Generic.List<char>();
        int remaining = number;
        char baseChar = upper ? 'A' : 'a';
        while (remaining > 0) {
            remaining--;
            chars.Add((char)(baseChar + (remaining % 26)));
            remaining /= 26;
        }

        chars.Reverse();
        return new string(chars.ToArray());
    }

}
