using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFooter(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont footerFont, string footerFontResource) {
        string text;
        var footerSegments = opts.GetFooterSegmentsForPage(variantPage);
        var footerZones = opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(footerZones)) {
            return BuildPageTextZones(opts, footerZones, page, pages, documentPages, footerFont, footerFontResource, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (footerSegments != null && footerSegments.Count > 0) {
            text = BuildPageTextFromSegments(footerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetFooterFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
        }
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double textWidth = EstimateSimpleTextWidthForOptions(text, footerFont, opts.FooterFontSize, opts);
        double x = opts.MarginLeft;
        if (opts.FooterAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.FooterAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.MarginBottom - opts.FooterOffsetY;
        PdfColor? footerColor = opts.FooterTextColor;
        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(footerFontResource, opts.FooterFontSize);
        if (footerColor.HasValue) {
            content.FillColor(footerColor.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeTextHex(text, footerFont, opts))
            .EndText();
        return sb.ToString();
    }

    private static string BuildHeader(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont headerFont, string headerFontResource) {
        string text;
        var headerSegments = opts.GetHeaderSegmentsForPage(variantPage);
        var headerZones = opts.GetHeaderZonesForPage(variantPage);
        if (HasPageTextZones(headerZones)) {
            return BuildPageTextZones(opts, headerZones, page, pages, documentPages, headerFont, headerFontResource, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (headerSegments != null && headerSegments.Count > 0) {
            text = BuildPageTextFromSegments(headerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetHeaderFormatForPage(variantPage), page, pages, documentPages, opts.PageNumberStyle);
        }

        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double textWidth = EstimateSimpleTextWidthForOptions(text, headerFont, opts.HeaderFontSize, opts);
        double x = opts.MarginLeft;
        if (opts.HeaderAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.HeaderAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY;
        PdfColor? headerColor = opts.HeaderTextColor;

        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(headerFontResource, opts.HeaderFontSize);
        if (headerColor.HasValue) {
            content.FillColor(headerColor.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeTextHex(text, headerFont, opts))
            .EndText();
        return sb.ToString();
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
        string fontResource,
        double fontSize,
        PdfColor? color,
        double offset,
        bool isHeader) {
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double y = isHeader ? opts.PageHeight - opts.MarginTop + offset : opts.MarginBottom - offset;
        var sb = new StringBuilder();
        var zoneLayouts = BuildPageTextZoneLayouts(opts, zones, page, pages, documentPages, font, fontSize, isHeader);
        foreach (var zone in zoneLayouts) {
            AppendPageText(sb, zone.Text, font, fontResource, fontSize, color, zone.X, y, opts);
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
            double textWidth = EstimateSimpleTextWidthForOptions(text, font, fontSize, opts);
            layouts.Add(("left", text, contentLeft, textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            string text = FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = EstimateSimpleTextWidthForOptions(text, font, fontSize, opts);
            layouts.Add(("center", text, contentLeft + ((contentWidth - textWidth) / 2), textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            string text = FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle);
            double textWidth = EstimateSimpleTextWidthForOptions(text, font, fontSize, opts);
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
