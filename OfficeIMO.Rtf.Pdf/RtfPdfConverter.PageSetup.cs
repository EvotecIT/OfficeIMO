using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void ApplyMetadata(RtfDocument document, PdfCore.PdfDocument pdf, RtfToPdfOptions options) {
        if (!options.IncludeMetadata) {
            return;
        }

        pdf.Meta(
            title: document.Info.Title,
            author: document.Info.Author,
            subject: document.Info.Subject,
            keywords: document.Info.Keywords);
    }

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfCreateOptions options) {
        if (setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) {
            options.PageWidth = RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value);
        }

        if (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) {
            options.PageHeight = RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value);
        }

        if ((setup.DirectLandscape == true && options.PageWidth < options.PageHeight) ||
            (setup.DirectLandscape == false && options.PageWidth > options.PageHeight)) {
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

    private static void ApplyPageSetup(RtfDocument document, RtfSection section, PdfCore.PdfPageBuilder page, PdfCore.PdfCreateOptions inheritedOptions) {
        ApplyPageSetup(document, section, page.Options, inheritedOptions);
    }

    private static void ApplyPageSetup(RtfDocument document, RtfSection section, PdfCore.PdfCreateOptions options, PdfCore.PdfCreateOptions inheritedOptions) {
        RtfPageSetup setup = document.GetPageSetup(section);
        options.PageWidth = inheritedOptions.PageWidth;
        options.PageHeight = inheritedOptions.PageHeight;
        options.MarginLeft = inheritedOptions.MarginLeft;
        options.MarginRight = inheritedOptions.MarginRight;
        options.MarginTop = inheritedOptions.MarginTop;
        options.MarginBottom = inheritedOptions.MarginBottom;
        options.PageBorder = inheritedOptions.PageBorder;
        options.PageNumberStyle = inheritedOptions.PageNumberStyle;
        ApplyPageSetup(document, setup, options);
        options.ClearPageNumberStartOverride();
        if (section.PageSetup.PageNumberRestart == true ||
            (section.PageSetup.PageNumberStart.HasValue && section.PageSetup.PageNumberRestart != false)) {
            options.PageNumberStart = section.PageSetup.PageNumberStart ?? document.PageSetup.PageNumberStart ?? 1;
        }
        options.PageStartParity = section.BreakKind == RtfSectionBreakKind.EvenPage ? PdfCore.PdfPageParity.Even
            : section.BreakKind == RtfSectionBreakKind.OddPage ? PdfCore.PdfPageParity.Odd
            : document.Settings.FacingPages == true && options.HasExplicitPageNumberStart
                ? options.PageNumberStart % 2 == 0 ? PdfCore.PdfPageParity.Even : PdfCore.PdfPageParity.Odd
                : null;
    }

    private static void ApplyHeaderFooters(RtfDocument document, RtfPageSetup setup, PdfCore.PdfCreateOptions options, RtfToPdfOptions saveOptions, IReadOnlyList<RtfHeaderFooter> declarations) {
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

        if (declarations.Count == 0) return;

        RecordHeaderFooterFrameDiagnostics(document, setup, declarations, saveOptions);

        string? defaultHeader = GetHeaderFooterText(declarations, RtfHeaderFooterKind.RightHeader)
            ?? GetHeaderFooterText(declarations, RtfHeaderFooterKind.Header);
        if (defaultHeader != null) {
            options.ShowHeader = defaultHeader.Length > 0;
            options.HeaderFormat = defaultHeader;
        }

        string? defaultFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.RightFooter)
            ?? GetHeaderFooterText(declarations, RtfHeaderFooterKind.Footer);
        if (defaultFooter != null) {
            options.ShowPageNumbers = defaultFooter.Length > 0;
            options.FooterFormat = defaultFooter;
        }

        string? firstHeader = GetHeaderFooterText(declarations, RtfHeaderFooterKind.FirstHeader);
        string? firstFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.FirstFooter);
        options.DifferentFirstPageHeaderFooter = setup.DifferentFirstPageHeaderFooter;
        options.FirstPageHeaderFormat = firstHeader ?? string.Empty;
        options.FirstPageFooterFormat = firstFooter ?? string.Empty;

        string? evenHeader = GetHeaderFooterText(declarations, RtfHeaderFooterKind.LeftHeader);
        string? evenFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.LeftFooter);
        options.DifferentOddAndEvenPagesHeaderFooter = document.Settings.FacingPages == true;
        options.UsePageNumberParityForHeaderFooter = true;
        options.EvenPageHeaderFormat = evenHeader ?? string.Empty;
        options.EvenPageFooterFormat = evenFooter ?? string.Empty;
    }

    private static string? GetHeaderFooterText(IReadOnlyList<RtfHeaderFooter> declarations, RtfHeaderFooterKind kind) {
        RtfHeaderFooter? headerFooter = GetHeaderFooter(declarations, kind);
        if (headerFooter == null) {
            return null;
        }

        string text = NormalizeHeaderFooterText(headerFooter.ToPlainText());
        return text;
    }

    private static void ApplySectionHeaderFooters(RtfDocument document, RtfSection section, PdfCore.PdfPageBuilder page,
        PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        if (!options.IncludeHeaderFooters) return;
        IReadOnlyList<RtfHeaderFooter> declarations = document.GetEffectiveHeaderFooters(section);
        if (declarations.Count == 0) return;
        RtfPageSetup setup = document.GetPageSetup(section);
        ApplyHeaderFooters(document, setup, page.Options, options, declarations);
        ApplyHeaderFooterContent(document, setup, pdf, page.Options, options, state, declarations,
            build => page.Header(build), build => page.Footer(build));
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
