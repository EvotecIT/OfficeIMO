using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string BuildFooter(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont footerFont, string footerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources) {
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs;
        var footerSegments = opts.GetFooterSegmentsForPage(variantPage);
        var footerZones = opts.GetFooterZonesForPage(variantPage);
        var footerZoneSegments = opts.GetFooterZoneSegmentsForPage(variantPage);
        if (footerZoneSegments?.HasContent == true) {
            return BuildPageTextSegmentZones(opts, footerZoneSegments, variantPage, page, pages, documentPages, footerFont, fontResources, namedFontResources, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (HasPageTextZones(footerZones)) {
            return BuildPageTextZones(opts, footerZones, variantPage, page, pages, documentPages, footerFont, fontResources, namedFontResources, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (footerSegments != null && footerSegments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(footerSegments, page, pages, documentPages, footerFont, opts.FooterFontSize, opts.FooterTextColor, opts, opts.FooterFontFamily);
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
        if (textWidth > 0D && TryGetPageTextVerticalBounds(runs, footerFont, opts.FooterFontSize, opts, y, out double textBottom, out double textTop)) {
            ReportHeaderFooterBounds(opts, "Footer", "text", x, textBottom, textWidth, textTop - textBottom);
        }
        PdfColor? footerColor = opts.FooterTextColor;
        var sb = new StringBuilder();
        AppendPageTextRuns(sb, runs, footerFont, footerFontResource, fontResources, namedFontResources, opts.FooterFontSize, footerColor, x, y, opts, textWidth, opts.FooterAlign);
        return sb.ToString();
    }

    private static string BuildHeader(PdfOptions opts, int variantPage, int page, int pages, int documentPages, PdfStandardFont headerFont, string headerFontResource, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, System.Collections.Generic.IReadOnlyDictionary<PdfNamedFontFace, string> namedFontResources) {
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs;
        var headerSegments = opts.GetHeaderSegmentsForPage(variantPage);
        var headerZones = opts.GetHeaderZonesForPage(variantPage);
        var headerZoneSegments = opts.GetHeaderZoneSegmentsForPage(variantPage);
        if (headerZoneSegments?.HasContent == true) {
            return BuildPageTextSegmentZones(opts, headerZoneSegments, variantPage, page, pages, documentPages, headerFont, fontResources, namedFontResources, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (HasPageTextZones(headerZones)) {
            return BuildPageTextZones(opts, headerZones, variantPage, page, pages, documentPages, headerFont, fontResources, namedFontResources, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (headerSegments != null && headerSegments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(headerSegments, page, pages, documentPages, headerFont, opts.HeaderFontSize, opts.HeaderTextColor, opts, opts.HeaderFontFamily);
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
        if (textWidth > 0D && TryGetPageTextVerticalBounds(runs, headerFont, opts.HeaderFontSize, opts, y, out double textBottom, out double textTop)) {
            ReportHeaderFooterBounds(opts, "Header", "text", x, textBottom, textWidth, textTop - textBottom);
        }
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
        PdfPageTextZoneSegments? zoneSegments = isHeader ? opts.GetHeaderZoneSegmentsForPage(variantPage) : opts.GetFooterZoneSegmentsForPage(variantPage);
        if (zoneSegments?.HasContent == true) {
            EnsurePageTextZoneSegmentFontResources(opts, zoneSegments, page, pages, documentPages, font, fontSize, fontFamily, ensureFontResource, ensureNamedFontResource);
            return;
        }
        var zones = isHeader ? opts.GetHeaderZonesForPage(variantPage) : opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(zones)) {
            EnsurePageTextZoneFontResources(opts, zones, page, pages, documentPages, font, fontSize, fontFamily, ensureFontResource, ensureNamedFontResource);
            return;
        }

        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs;
        var segments = isHeader ? opts.GetHeaderSegmentsForPage(variantPage) : opts.GetFooterSegmentsForPage(variantPage);
        if (segments != null && segments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(segments, page, pages, documentPages, font, fontSize, color: null, opts, fontFamily);
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

    private static void EnsurePageTextZoneSegmentFontResources(
        PdfOptions opts,
        PdfPageTextZoneSegments zones,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        string? fontFamily,
        Func<PdfStandardFont, string, string> ensureFontResource,
        Action<PdfNamedFontFace> ensureNamedFontResource) {
        foreach (System.Collections.Generic.IReadOnlyList<FooterSegment>? segments in new[] { zones.Left, zones.Center, zones.Right }) {
            if (segments == null || segments.Count == 0) continue;
            EnsurePageTextRunFontResources(
                BuildPageTextRunsFromSegments(segments, page, pages, documentPages, font, fontSize, color: null, opts, fontFamily),
                font, opts, ensureFontResource, ensureNamedFontResource);
        }
    }

    private static void EnsurePageTextRunFontResources(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont, PdfOptions opts, Func<PdfStandardFont, string, string> ensureFontResource, Action<PdfNamedFontFace> ensureNamedFontResource) {
        PdfStandardFont normalFont = ChooseNormal(opts.DefaultFont);
        foreach (PdfTextRun run in runs) {
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
            if (string.IsNullOrEmpty(zone.Text)) {
                continue;
            }

            string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildPageTextRuns(zone.Text, font, fontSize, color, opts, fontFamily);
            PdfNamedFontFace? namedFont = TryResolvePageTextNamedFont(opts, fontFamily, font, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            string baseFontResource = ResolvePageTextFontResource(fontResources, namedFontResources, font, namedFont);
            AppendPageTextRuns(sb, runs, font, baseFontResource, fontResources, namedFontResources, fontSize, color, zone.X, y, opts, zone.TextWidth, zone.Align);
        }

        return sb.ToString();
    }

    private static string BuildPageTextSegmentZones(
        PdfOptions opts,
        PdfPageTextZoneSegments zones,
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
        foreach (PageTextRichZoneLayout zone in BuildPageTextSegmentZoneLayouts(opts, zones, variantPage, page, pages, documentPages, font, fontSize, color, isHeader)) {
            string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;
            PdfNamedFontFace? namedFont = TryResolvePageTextNamedFont(opts, fontFamily, font, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            string baseFontResource = ResolvePageTextFontResource(fontResources, namedFontResources, font, namedFont);
            AppendPageTextRuns(sb, zone.Runs, font, baseFontResource, fontResources, namedFontResources, fontSize, color, zone.X, y, opts, zone.TextWidth, zone.Align);
        }
        return sb.ToString();
    }

    private static System.Collections.Generic.List<PageTextRichZoneLayout> BuildPageTextSegmentZoneLayouts(
        PdfOptions opts,
        PdfPageTextZoneSegments zones,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfStandardFont font,
        double fontSize,
        PdfColor? color,
        bool isHeader) {
        double contentLeft = opts.MarginLeft;
        double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double textBaseline = isHeader ? opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY : opts.MarginBottom - opts.FooterOffsetY;
        var layouts = new System.Collections.Generic.List<PageTextRichZoneLayout>();
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images = isHeader ? opts.GetHeaderImagesForPage(variantPage) : opts.GetFooterImagesForPage(variantPage);
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes = isHeader ? opts.GetHeaderShapesForPage(variantPage) : opts.GetFooterShapesForPage(variantPage);
        string? fontFamily = isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily;

        Add(PdfAlign.Left, zones.Left);
        Add(PdfAlign.Center, zones.Center);
        Add(PdfAlign.Right, zones.Right);
        ValidatePageTextRichZoneLayouts(opts, layouts, isHeader);
        return layouts;

        void Add(PdfAlign align, System.Collections.Generic.IReadOnlyList<FooterSegment>? segments) {
            if (segments == null || segments.Count == 0) return;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildPageTextRunsFromSegments(segments, page, pages, documentPages, font, fontSize, color, opts, fontFamily);
            double textWidth = MeasurePageTextRuns(runs, font, fontSize, opts);
            GetPageTextVerticalBoundsOrBaseline(runs, font, fontSize, opts, textBaseline, out double textBottom, out double textTop);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, MeasureHeaderFooterImagesWidth(images, align), MeasureHeaderFooterShapesWidth(shapes, align));
            double x = align == PdfAlign.Center
                ? contentLeft + ((contentWidth - occupiedWidth) / 2D)
                : align == PdfAlign.Right ? contentLeft + contentWidth - occupiedWidth : contentLeft;
            layouts.Add(new PageTextRichZoneLayout(runs, x, textWidth, align, textBottom, textTop - textBottom));
        }
    }

    private static void ValidatePageTextRichZoneLayouts(PdfOptions options, System.Collections.Generic.List<PageTextRichZoneLayout> layouts, bool isHeader) {
        string source = isHeader ? "Header" : "Footer";
        foreach (PageTextRichZoneLayout zone in layouts) {
            if (zone.TextWidth > 0D) ReportHeaderFooterBounds(options, source, "text", zone.X, zone.TextBottom, zone.TextWidth, zone.TextHeight);
        }
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
        double textBaseline = isHeader
            ? opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY
            : opts.MarginBottom - opts.FooterOffsetY;
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
            System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily);
            double textWidth = MeasurePageTextRuns(runs, font, fontSize, opts);
            GetPageTextVerticalBoundsOrBaseline(runs, font, fontSize, opts, textBaseline, out double textBottom, out double textTop);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Left);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Left);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            layouts.Add(new PageTextZoneLayout(text, contentLeft, textWidth, PdfAlign.Left, textBottom, textTop - textBottom));
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            string text = FormatPageText(zones.Center!, page, pages, documentPages, opts.PageNumberStyle);
            System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily);
            double textWidth = MeasurePageTextRuns(runs, font, fontSize, opts);
            GetPageTextVerticalBoundsOrBaseline(runs, font, fontSize, opts, textBaseline, out double textBottom, out double textTop);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Center);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Center);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            double occupiedX = contentLeft + ((contentWidth - occupiedWidth) / 2);
            layouts.Add(new PageTextZoneLayout(text, occupiedX, textWidth, PdfAlign.Center, textBottom, textTop - textBottom));
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            string text = FormatPageText(zones.Right!, page, pages, documentPages, opts.PageNumberStyle);
            System.Collections.Generic.IReadOnlyList<PdfTextRun> runs = BuildPageTextRuns(text, font, fontSize, color: null, opts, fontFamily);
            double textWidth = MeasurePageTextRuns(runs, font, fontSize, opts);
            GetPageTextVerticalBoundsOrBaseline(runs, font, fontSize, opts, textBaseline, out double textBottom, out double textTop);
            double imagesWidth = MeasureHeaderFooterImagesWidth(images, PdfAlign.Right);
            double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, PdfAlign.Right);
            double occupiedWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
            double occupiedX = contentLeft + contentWidth - occupiedWidth;
            layouts.Add(new PageTextZoneLayout(text, occupiedX, textWidth, PdfAlign.Right, textBottom, textTop - textBottom));
        }

        ValidatePageTextZoneLayouts(opts, layouts, isHeader);
        return layouts;
    }

    private static void ValidatePageTextZoneLayouts(PdfOptions options, System.Collections.Generic.List<PageTextZoneLayout> layouts, bool isHeader) {
        string source = isHeader ? "Header" : "Footer";
        foreach (var zone in layouts) {
            if (zone.TextWidth > 0D) {
                ReportHeaderFooterBounds(options, source, "text", zone.X, zone.TextBottom, zone.TextWidth, zone.TextHeight);
            }
        }
    }

    private static void ValidateHeaderFooterGroupLayouts(
        PdfOptions options,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        bool isHeader) {
        const double tolerance = 0.01D;
        const double minimumGap = 2D;
        double contentLeft = options.MarginLeft;
        double contentRight = options.PageWidth - options.MarginRight;
        string source = isHeader ? "Header" : "Footer";
        var layouts = new System.Collections.Generic.List<HeaderFooterGroupLayout>(3);

        foreach (PdfAlign align in new[] { PdfAlign.Left, PdfAlign.Center, PdfAlign.Right }) {
            if (!TryGetHeaderFooterGroupVisibleHorizontalBounds(
                    options,
                    variantPage,
                    page,
                    pages,
                    documentPages,
                    align,
                    isHeader,
                    out double visibleLeft,
                    out double visibleRight)) {
                continue;
            }

            var layout = new HeaderFooterGroupLayout(visibleLeft, visibleRight - visibleLeft);
            layouts.Add(layout);
            if (layout.X < contentLeft - tolerance || layout.X + layout.Width > contentRight + tolerance) {
                options.AddLayoutDiagnostic(
                    "HeaderFooterContentOverflow",
                    source,
                    source + " zone content exceeds the page content frame and was preserved as authored.",
                    PdfLayoutDiagnosticKind.Overflow,
                    PdfConversionWarningSeverity.Information,
                    x: layout.X,
                    width: layout.Width);
            }
        }

        var ordered = layouts.OrderBy(zone => zone.X).ToList();
        for (int i = 1; i < ordered.Count; i++) {
            var previous = ordered[i - 1];
            var current = ordered[i];
            if (previous.X + previous.Width + minimumGap > current.X + tolerance) {
                options.AddLayoutDiagnostic(
                    "HeaderFooterZoneOverlap",
                    source,
                    source + " zones overlap and were preserved as authored.",
                    PdfLayoutDiagnosticKind.Overflow,
                    PdfConversionWarningSeverity.Information);
            }
        }
    }

    private static bool TryGetHeaderFooterGroupVisibleHorizontalBounds(
        PdfOptions options,
        int variantPage,
        int page,
        int pages,
        int documentPages,
        PdfAlign align,
        bool isHeader,
        out double visibleLeft,
        out double visibleRight) {
        double textWidth = MeasureHeaderFooterTextWidth(options, variantPage, page, pages, documentPages, align, isHeader);
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> images = isHeader
            ? options.GetHeaderImagesForPage(variantPage)
            : options.GetFooterImagesForPage(variantPage);
        System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> shapes = isHeader
            ? options.GetHeaderShapesForPage(variantPage)
            : options.GetFooterShapesForPage(variantPage);
        double imagesWidth = MeasureHeaderFooterImagesWidth(images, align);
        double shapesWidth = MeasureHeaderFooterShapesWidth(shapes, align);
        double groupWidth = CombineHeaderFooterInlineWidths(textWidth, imagesWidth, shapesWidth);
        double groupX = AlignHeaderFooterGroup(options, groupWidth, align);
        visibleLeft = double.PositiveInfinity;
        visibleRight = double.NegativeInfinity;

        if (textWidth > 0D) {
            IncludeHorizontalBounds(groupX, groupX + textWidth, ref visibleLeft, ref visibleRight);
        }

        double imageConsumedWidth = 0D;
        foreach (PdfHeaderFooterImage image in images) {
            if (image.Align != align) {
                continue;
            }

            double imageX = groupX +
                (textWidth > 0D ? textWidth + HeaderFooterInlineGap : 0D) +
                imageConsumedWidth;
            double imageBottomY = isHeader
                ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - image.Height
                : options.MarginBottom - options.FooterOffsetY;
            ImageBlock block = image.ToImageBlock();
            PdfImageStyle style = block.Style ?? new PdfImageStyle();
            PageImage pageImage = CreatePageImage(block, style, imageX, imageBottomY);
            GetImageAnnotationBounds(pageImage, out double imageLeft, out _, out double imageRight, out _);
            IncludeHorizontalBounds(imageLeft, imageRight, ref visibleLeft, ref visibleRight);
            imageConsumedWidth += image.Width + HeaderFooterInlineGap;
        }

        double shapeConsumedWidth = 0D;
        foreach (PdfHeaderFooterShape headerFooterShape in shapes) {
            if (headerFooterShape.Align != align) {
                continue;
            }

            double shapeX = groupX +
                (textWidth > 0D ? textWidth + HeaderFooterInlineGap : 0D) +
                (imagesWidth > 0D ? imagesWidth + HeaderFooterInlineGap : 0D) +
                shapeConsumedWidth;
            ShapeBlock shapeBlock = headerFooterShape.ToShapeBlock();
            double shapeBottomY = isHeader
                ? options.PageHeight - options.MarginTop + options.HeaderOffsetY - shapeBlock.Shape.Height
                : options.MarginBottom - options.FooterOffsetY;
            if (TryGetHeaderFooterShapeBounds(shapeBlock.Shape, shapeX, shapeBottomY, out double shapeLeft, out _, out double shapeWidth, out _)) {
                IncludeHorizontalBounds(shapeLeft, shapeLeft + shapeWidth, ref visibleLeft, ref visibleRight);
            }
            shapeConsumedWidth += headerFooterShape.Width + HeaderFooterInlineGap;
        }

        return !double.IsPositiveInfinity(visibleLeft) && !double.IsNegativeInfinity(visibleRight);
    }

    private static void IncludeHorizontalBounds(double left, double right, ref double visibleLeft, ref double visibleRight) {
        visibleLeft = System.Math.Min(visibleLeft, left);
        visibleRight = System.Math.Max(visibleRight, right);
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
        PdfPageTextZoneSegments? zoneSegments = isHeader ? opts.GetHeaderZoneSegmentsForPage(variantPage) : opts.GetFooterZoneSegmentsForPage(variantPage);
        System.Collections.Generic.List<FooterSegment>? alignedZoneSegments = align switch {
            PdfAlign.Center => zoneSegments?.Center,
            PdfAlign.Right => zoneSegments?.Right,
            _ => zoneSegments?.Left
        };
        string? text = align switch {
            PdfAlign.Center => zones.Center,
            PdfAlign.Right => zones.Right,
            _ => zones.Left
        };
        System.Collections.Generic.IReadOnlyList<PdfTextRun>? runs = null;

        if (alignedZoneSegments != null && alignedZoneSegments.Count > 0) {
            runs = BuildPageTextRunsFromSegments(alignedZoneSegments, page, pages, documentPages, font, fontSize, color: null, opts, isHeader ? opts.HeaderFontFamily : opts.FooterFontFamily);
        } else if (!string.IsNullOrEmpty(text)) {
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
                    documentPages,
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
        double TextBottom,
        double TextHeight);

    private readonly record struct PageTextRichZoneLayout(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs,
        double X,
        double TextWidth,
        PdfAlign Align,
        double TextBottom,
        double TextHeight);

    private readonly record struct HeaderFooterGroupLayout(double X, double Width);

}
