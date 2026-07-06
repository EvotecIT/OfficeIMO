using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void ApplyMetadata(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        if (!options.IncludeMetadata) {
            return;
        }

        pdf.Meta(
            title: document.Info.Title,
            author: document.Info.Author,
            subject: document.Info.Subject,
            keywords: document.Info.Keywords);
    }

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfOptions options) {
        if (setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) {
            options.PageWidth = RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value);
        }

        if (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) {
            options.PageHeight = RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value);
        }

        if (setup.Landscape && options.PageWidth < options.PageHeight) {
            double width = options.PageWidth;
            options.PageWidth = options.PageHeight;
            options.PageHeight = width;
        }

        if (setup.MarginLeftTwips.HasValue) {
            options.MarginLeft = RtfPdfMapping.TwipsToPoints(setup.MarginLeftTwips.Value);
        }

        if (setup.MarginRightTwips.HasValue) {
            options.MarginRight = RtfPdfMapping.TwipsToPoints(setup.MarginRightTwips.Value);
        }

        if (setup.MarginTopTwips.HasValue) {
            options.MarginTop = RtfPdfMapping.TwipsToPoints(setup.MarginTopTwips.Value);
        }

        if (setup.MarginBottomTwips.HasValue) {
            options.MarginBottom = RtfPdfMapping.TwipsToPoints(setup.MarginBottomTwips.Value);
        }

        if (setup.PageNumberStart.HasValue) {
            options.PageNumberStart = setup.PageNumberStart.Value;
        }

        if (setup.PageNumberFormat.HasValue) {
            options.PageNumberStyle = RtfPdfMapping.ToPdfPageNumberStyle(setup.PageNumberFormat.Value);
        }

        PdfCore.PdfPageBorder? border = RtfPdfMapping.ToPdfPageBorder(document, setup.PageBorders);
        if (border != null) {
            options.PageBorder = border;
        }
    }

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfPageCompose page, PdfCore.PdfOptions inheritedOptions) {
        double width = setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0
            ? RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value)
            : inheritedOptions.PageWidth;
        double height = setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0
            ? RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value)
            : inheritedOptions.PageHeight;

        if (setup.Landscape && width < height) {
            double swap = width;
            width = height;
            height = swap;
        }

        if ((setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) ||
            (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) ||
            setup.Landscape) {
            page.Size(width, height);
        }

        if (HasAnyMargin(setup)) {
            page.Margin(
                setup.MarginLeftTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginLeftTwips.Value) : inheritedOptions.MarginLeft,
                setup.MarginTopTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginTopTwips.Value) : inheritedOptions.MarginTop,
                setup.MarginRightTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginRightTwips.Value) : inheritedOptions.MarginRight,
                setup.MarginBottomTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginBottomTwips.Value) : inheritedOptions.MarginBottom);
        }

        if (setup.PageNumberStart.HasValue) {
            page.PageNumberStart(setup.PageNumberStart.Value);
        }

        if (setup.PageNumberFormat.HasValue) {
            page.PageNumberStyle(RtfPdfMapping.ToPdfPageNumberStyle(setup.PageNumberFormat.Value));
        }

        PdfCore.PdfPageBorder? border = RtfPdfMapping.ToPdfPageBorder(document, setup.PageBorders);
        if (border != null) {
            page.PageBorder(border);
        }
    }

    private static bool HasAnyMargin(RtfPageSetup setup) {
        return setup.MarginLeftTwips.HasValue ||
               setup.MarginRightTwips.HasValue ||
               setup.MarginTopTwips.HasValue ||
               setup.MarginBottomTwips.HasValue;
    }

    private static void ApplyHeaderFooters(RtfDocument document, PdfCore.PdfOptions options, RtfPdfSaveOptions saveOptions) {
        if (document.HeaderFooters.Count == 0) {
            return;
        }

        if (!saveOptions.IncludeHeaderFooters) {
            AddConversionWarning(
                saveOptions,
                "HeaderFooterSkipped",
                "HeaderFooter",
                "RTF header and footer text was skipped because IncludeHeaderFooters is false.",
                new Dictionary<string, string> {
                    ["Count"] = document.HeaderFooters.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)
                });
            return;
        }

        string? defaultHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.RightHeader)
            ?? GetHeaderFooterText(document, RtfHeaderFooterKind.Header);
        if (defaultHeader != null && defaultHeader.Length > 0) {
            options.ShowHeader = true;
            options.HeaderFormat = defaultHeader;
        }

        string? defaultFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.RightFooter)
            ?? GetHeaderFooterText(document, RtfHeaderFooterKind.Footer);
        if (defaultFooter != null && defaultFooter.Length > 0) {
            options.ShowPageNumbers = true;
            options.FooterFormat = defaultFooter;
        }

        string? firstHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.FirstHeader);
        string? firstFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.FirstFooter);
        if ((firstHeader != null && firstHeader.Length > 0) ||
            (firstFooter != null && firstFooter.Length > 0) ||
            document.PageSetup.DifferentFirstPageHeaderFooter) {
            options.DifferentFirstPageHeaderFooter = true;
            if (firstHeader != null && firstHeader.Length > 0) {
                options.FirstPageHeaderFormat = firstHeader;
            }

            if (firstFooter != null && firstFooter.Length > 0) {
                options.FirstPageFooterFormat = firstFooter;
            }
        }

        string? evenHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.LeftHeader);
        string? evenFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.LeftFooter);
        if ((evenHeader != null && evenHeader.Length > 0) ||
            (evenFooter != null && evenFooter.Length > 0)) {
            options.DifferentOddAndEvenPagesHeaderFooter = true;
            if (evenHeader != null && evenHeader.Length > 0) {
                options.EvenPageHeaderFormat = evenHeader;
            }

            if (evenFooter != null && evenFooter.Length > 0) {
                options.EvenPageFooterFormat = evenFooter;
            }
        }
    }

    private static string? GetHeaderFooterText(RtfDocument document, RtfHeaderFooterKind kind) {
        RtfHeaderFooter? headerFooter = document.HeaderFooters.FirstOrDefault(item => item.Kind == kind);
        if (headerFooter == null) {
            return null;
        }

        string text = NormalizeHeaderFooterText(headerFooter.ToPlainText());
        return text.Length == 0 ? null : text;
    }

    private static string NormalizeHeaderFooterText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return text
            .Replace("\r\n", " ")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\f', ' ')
            .Replace('\v', ' ')
            .Trim();
    }
}
