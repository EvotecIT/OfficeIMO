using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFooter(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont footerFont, string footerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources) {
        string text;
        var footerSegments = opts.GetFooterSegmentsForPage(variantPage);
        var footerZones = opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(footerZones)) {
            return BuildPageTextZones(opts, footerZones, page, pages, documentPages, footerFont, fontResources, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (footerSegments != null && footerSegments.Count > 0) {
            text = BuildPageTextFromSegments(footerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetFooterFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
        }
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildPageTextRuns(text, footerFont, opts.FooterFontSize, opts.FooterTextColor, opts);
        double textWidth = MeasurePageTextRuns(runs, footerFont, opts.FooterFontSize, opts);
        double x = opts.MarginLeft;
        if (opts.FooterAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.FooterAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.MarginBottom - opts.FooterOffsetY;
        PdfColor? footerColor = opts.FooterTextColor;
        var sb = new StringBuilder();
        AppendPageTextRuns(sb, runs, footerFont, footerFontResource, fontResources, opts.FooterFontSize, footerColor, x, y, opts);
        return sb.ToString();
    }

    private static string BuildHeader(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont headerFont, string headerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources) {
        string text;
        var headerSegments = opts.GetHeaderSegmentsForPage(variantPage);
        var headerZones = opts.GetHeaderZonesForPage(variantPage);
        if (HasPageTextZones(headerZones)) {
            return BuildPageTextZones(opts, headerZones, page, pages, documentPages, headerFont, fontResources, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (headerSegments != null && headerSegments.Count > 0) {
            text = BuildPageTextFromSegments(headerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetHeaderFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
        }

        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildPageTextRuns(text, headerFont, opts.HeaderFontSize, opts.HeaderTextColor, opts);
        double textWidth = MeasurePageTextRuns(runs, headerFont, opts.HeaderFontSize, opts);
        double x = opts.MarginLeft;
        if (opts.HeaderAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.HeaderAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY;
        PdfColor? headerColor = opts.HeaderTextColor;

        var sb = new StringBuilder();
        AppendPageTextRuns(sb, runs, headerFont, headerFontResource, fontResources, opts.HeaderFontSize, headerColor, x, y, opts);
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
        Func<PdfStandardFont, string, string> ensureFontResource) {
        var zones = isHeader ? opts.GetHeaderZonesForPage(variantPage) : opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(zones)) {
            EnsurePageTextZoneFontResources(opts, zones, page, pages, documentPages, font, fontSize, ensureFontResource);
            return;
        }

        string text;
        var segments = isHeader ? opts.GetHeaderSegmentsForPage(variantPage) : opts.GetFooterSegmentsForPage(variantPage);
        if (segments != null && segments.Count > 0) {
            text = BuildPageTextFromSegments(segments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(isHeader ? opts.GetHeaderFormatForPage(variantPage) : opts.GetFooterFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
        }

        EnsurePageTextRunFontResources(BuildPageTextRuns(text, font, fontSize, color: null, opts), font, opts, ensureFontResource);
    }

    private static void EnsurePageTextZoneFontResources(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        Func<PdfStandardFont, string, string> ensureFontResource) {
        if (!string.IsNullOrEmpty(zones.Left)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Left!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts), font, opts, ensureFontResource);
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts), font, opts, ensureFontResource);
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            EnsurePageTextRunFontResources(BuildPageTextRuns(FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle), font, fontSize, color: null, opts), font, opts, ensureFontResource);
        }
    }

    private static void EnsurePageTextRunFontResources(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, PdfOptions opts, Func<PdfStandardFont, string, string> ensureFontResource) {
        PdfStandardFont normalFont = ChooseNormal(opts.DefaultFont);
        foreach (TextRun run in runs) {
            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            ensureFontResource(runFont, GetStandardFontResourceName(runFont, normalFont));
        }
    }

    private static bool HasPageTextZones((string? Left, string? Center, string? Right) zones) =>
        !string.IsNullOrEmpty(zones.Left) ||
        !string.IsNullOrEmpty(zones.Center) ||
        !string.IsNullOrEmpty(zones.Right);

    private static string BuildPageTextZones(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources,
        double fontSize,
        PdfColor? color,
        double offset,
        bool isHeader) {
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double y = isHeader ? opts.PageHeight - opts.MarginTop + offset : opts.MarginBottom - offset;
        var sb = new StringBuilder();
        var zoneLayouts = BuildPageTextZoneLayouts(opts, zones, page, pages, documentPages, font, fontSize, isHeader);
        foreach (var zone in zoneLayouts) {
            System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildPageTextRuns(zone.Text, font, fontSize, color, opts);
            AppendPageTextRuns(sb, runs, font, ResolvePageTextFontResource(fontResources, font), fontResources, fontSize, color, zone.X, y, opts);
        }

        return sb.ToString();
    }

    private static System.Collections.Generic.List<(string Name, string Text, double X, double Width)> BuildPageTextZoneLayouts(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        bool isHeader) {
        double contentLeft = opts.MarginLeft;
        double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        var layouts = new System.Collections.Generic.List<(string Name, string Text, double X, double Width)>();

        if (!string.IsNullOrEmpty(zones.Left)) {
            string text = FormatPageText(zones.Left!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts), font, fontSize, opts);
            layouts.Add(("left", text, contentLeft, textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            string text = FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts), font, fontSize, opts);
            layouts.Add(("center", text, contentLeft + ((contentWidth - textWidth) / 2), textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            string text = FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = MeasurePageTextRuns(BuildPageTextRuns(text, font, fontSize, color: null, opts), font, fontSize, opts);
            layouts.Add(("right", text, contentLeft + contentWidth - textWidth, textWidth));
        }

        ValidatePageTextZoneLayouts(layouts, contentLeft, contentLeft + contentWidth, isHeader);
        return layouts;
    }

    private static void ValidatePageTextZoneLayouts(System.Collections.Generic.List<(string Name, string Text, double X, double Width)> layouts, double contentLeft, double contentRight, bool isHeader) {
        const double tolerance = 0.01D;
        const double minimumGap = 2D;
        string scope = isHeader ? "header" : "footer";
        foreach (var zone in layouts) {
            if (zone.X < contentLeft - tolerance || zone.X + zone.Width > contentRight + tolerance) {
                throw new ArgumentException("PDF " + scope + " zone text must fit inside the page content width.");
            }
        }

        var ordered = layouts.OrderBy(zone => zone.X).ToList();
        for (int i = 1; i < ordered.Count; i++) {
            var previous = ordered[i - 1];
            var current = ordered[i];
            if (previous.X + previous.Width + minimumGap > current.X + tolerance) {
                throw new ArgumentException("PDF " + scope + " zone text must not overlap.");
            }
        }
    }

    private static System.Collections.Generic.IReadOnlyList<TextRun> BuildPageTextRuns(string text, PdfStandardFont font, double fontSize, PdfColor? color, PdfOptions opts) =>
        NormalizeFallbackRuns(new[] { TextRun.Normal(text, color, fontSize, font: font) }, ChooseNormal(font), opts);

    private static double MeasurePageTextRuns(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions opts) {
        double width = 0D;
        foreach (TextRun run in runs) {
            width += MeasureRichText(run.Text ?? string.Empty, ResolvePageTextRunFont(run, baseFont), run.FontSize ?? fontSize, run.Baseline, opts);
        }

        return width;
    }

    private static void AppendPageTextRuns(
        StringBuilder sb,
        System.Collections.Generic.IReadOnlyList<TextRun> runs,
        PdfStandardFont baseFont,
        string baseFontResource,
        System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources,
        double fontSize,
        PdfColor? color,
        double x,
        double y,
        PdfOptions opts) {
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(baseFontResource, fontSize);
        if (color.HasValue) {
            content.FillColor(color.Value);
        }

        content.TextMatrix(x, y);
        foreach (TextRun run in runs) {
            string text = run.Text ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            string fontResource = ResolvePageTextFontResource(fontResources, runFont);
            double runFontSize = run.FontSize ?? fontSize;
            content
                .Font(fontResource, runFontSize)
                .ShowHexText(EncodeTextHex(text, runFont, opts));
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

    private static void AppendPageText(StringBuilder sb, string text, PdfStandardFont font, string fontResource, double fontSize, PdfColor? color, double x, double y, PdfOptions opts) {
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(fontResource, fontSize);
        if (color.HasValue) {
            content.FillColor(color.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeTextHex(text, font, opts))
            .EndText();
    }

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
