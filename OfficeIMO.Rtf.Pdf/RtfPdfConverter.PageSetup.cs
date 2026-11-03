using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void ApplyMetadata(RtfDocument document, PdfCore.PdfDocument pdf, RtfToPdfOptions options) {
        if (!options.IncludeMetadata) {
            return;
        }

        pdf.Meta(
            title: document.Properties.Title,
            author: document.Properties.Author,
            subject: document.Properties.Subject,
            keywords: document.Properties.Keywords);
    }

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfCreateOptions options) {
        if (setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) {
            options.PageSetup.Size = new global::OfficeIMO.Pdf.PdfPageSize(global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value)), options.PageSetup.Size.Height);
        }

        if (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) {
            options.PageSetup.Size = new global::OfficeIMO.Pdf.PdfPageSize(options.PageSetup.Size.Width, global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value)));
        }

        if (setup.DirectLandscape.HasValue) {
            options.PageSetup.Orientation = setup.DirectLandscape.Value ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
        }

        if (setup.MarginLeftTwips.HasValue) {
            options.PageSetup.Margins.Left = global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.MarginLeftTwips.Value));
        }

        if (setup.MarginRightTwips.HasValue) {
            options.PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.MarginRightTwips.Value));
        }

        if (setup.MarginTopTwips.HasValue) {
            options.PageSetup.Margins.Top = global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.MarginTopTwips.Value));
        }

        if (setup.MarginBottomTwips.HasValue) {
            options.PageSetup.Margins.Bottom = global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(setup.MarginBottomTwips.Value));
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
        options.PageSetup.Size = inheritedOptions.PageSetup.Size;
        options.PageSetup.Margins.Left = inheritedOptions.PageSetup.Margins.Left;
        options.PageSetup.Margins.Right = inheritedOptions.PageSetup.Margins.Right;
        options.PageSetup.Margins.Top = inheritedOptions.PageSetup.Margins.Top;
        options.PageSetup.Margins.Bottom = inheritedOptions.PageSetup.Margins.Bottom;
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
            options.PageSetup.Header.Show = defaultHeader.Length > 0;
            options.PageSetup.Header.Format = defaultHeader;
        }

        string? defaultFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.RightFooter)
            ?? GetHeaderFooterText(declarations, RtfHeaderFooterKind.Footer);
        if (defaultFooter != null) {
            options.PageSetup.Footer.Show = defaultFooter.Length > 0;
            options.PageSetup.Footer.Format = defaultFooter;
        }

        string? firstHeader = GetHeaderFooterText(declarations, RtfHeaderFooterKind.FirstHeader);
        string? firstFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.FirstFooter);
        options.PageSetup.DifferentFirstPage = setup.DifferentFirstPageHeaderFooter;
        options.PageSetup.FirstPageHeader.Format = firstHeader ?? string.Empty;
        options.PageSetup.FirstPageFooter.Format = firstFooter ?? string.Empty;

        string? evenHeader = GetHeaderFooterText(declarations, RtfHeaderFooterKind.LeftHeader);
        string? evenFooter = GetHeaderFooterText(declarations, RtfHeaderFooterKind.LeftFooter);
        options.PageSetup.DifferentOddAndEvenPages = document.Settings.FacingPages == true;
        options.PageSetup.UseVisiblePageNumberParity = true;
        options.PageSetup.EvenPageHeader.Format = evenHeader ?? string.Empty;
        options.PageSetup.EvenPageFooter.Format = evenFooter ?? string.Empty;
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
