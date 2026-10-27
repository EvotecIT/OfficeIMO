using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    internal static PdfCore.PdfDocument Convert(RtfDocument document, RtfToPdfOptions? options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        RtfTableTraversalGuard.ValidateDocument(document);
        RtfToPdfOptions normalized = options ?? new RtfToPdfOptions();
        foreach (RtfConversionDiagnostic diagnostic in document.GetStyleConversionDiagnostics().Diagnostics) {
            AddConversionWarning(normalized, diagnostic.Code, diagnostic.SourcePath ?? "Styles", diagnostic.Message, diagnostic.Action);
        }
        PdfCore.PdfCreateOptions pdfOptions = normalized.PdfOptions ?? new PdfCore.PdfCreateOptions();
        // RTF leaves pair kerning disabled unless a run supplies a positive threshold.
        pdfOptions.TextKerning = false;
        if (!pdfOptions.HasExplicitTextWhitespaceConfiguration) {
            pdfOptions.TextWhitespaceMode = PdfCore.PdfTextWhitespaceMode.Preserve;
        }
        pdfOptions.UseContentStreamCompressionByDefault();
        pdfOptions.ReportDiagnosticsTo(normalized.Report, "OfficeIMO.Rtf.Pdf");
        IReadOnlyDictionary<int, PdfCore.PdfStandardFont> fontSlots = ConfigureDocumentFonts(
            document,
            pdfOptions,
            normalized);
        ApplyPageSetup(document, document.PageSetup, pdfOptions);
        PdfCore.PdfCreateOptions documentDefaults = pdfOptions.Clone();
        RtfPageSetup firstPageSetup = document.PageSetup;
        if (document.Sections.Count > 0) {
            firstPageSetup = document.GetPageSetup(document.Sections[0]);
            ApplyPageSetup(document, firstPageSetup, pdfOptions);
        }

        ApplyHeaderFooters(document, firstPageSetup, pdfOptions, normalized,
            document.Sections.Count == 0 ? document.HeaderFooters : document.GetEffectiveHeaderFooters(document.Sections[0]));

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        ApplyMetadata(document, pdf, normalized);
        PdfRenderState state = new PdfRenderState(document, fontSlots, pdf.Options);
        ApplyHeaderFooterContent(document, firstPageSetup, pdf, pdf.Options, normalized, state,
            document.Sections.Count == 0 ? document.HeaderFooters : document.GetEffectiveHeaderFooters(document.Sections[0]),
            build => pdf.Header(build), build => pdf.Footer(build));

        RenderDocumentBlocks(document, pdf, normalized, state, documentDefaults);

        RenderNotes(document, pdf, normalized, state);
        return pdf;
    }

    private static void RenderDocumentBlocks(RtfDocument document, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state, PdfCore.PdfCreateOptions pdfOptions) {
        if (document.Sections.Count == 0) {
            RenderBlocks(document, document.Blocks, pdf, options, state);
            return;
        }

        PdfCore.PdfCreateOptions currentPageOptions = pdf.Options.Clone();
        RtfSection currentPageSection = document.Sections[0];
        options.CancellationToken.ThrowIfCancellationRequested();
        bool canContinueRunningStories = SupportsContinuousRunningStoryDocument(document);
        for (int index = 0; index < document.Sections.Count; index++) {
            RtfSection section = document.Sections[index];
            if (index == 0) {
                RenderSectionBlocks(document, section, pdf, options, state);
                continue;
            }

            if (!StartsNewPdfPage(section.BreakKind)) {
                RenderContinuousSectionBlocks(document, section, index, pdf, options, state, pdfOptions,
                    canContinueRunningStories, ref currentPageSection, ref currentPageOptions);
                continue;
            }

            if (section.BreakKind == RtfSectionBreakKind.Column &&
                ((section.ColumnCount ?? section.Columns.Count) > 1 ||
                 (document.Sections[index - 1].ColumnCount ?? document.Sections[index - 1].Columns.Count) > 1)) {
                AddConversionWarning(options, "ColumnSectionBreakFlattened", "Section/Break",
                    "An RTF column-starting section starts a new PDF page because continuation within the preceding column flow is not supported.",
                    RtfConversionAction.Flattened);
            }

            pdf.Section(page => {
                ApplyPageSetup(document, section, page, pdfOptions);
                ApplySectionHeaderFooters(document, section, page, pdf, options, state);
                currentPageOptions = page.Options.Clone();
                currentPageSection = section;
                PdfCore.PdfCreateOptions previousFontOptions = state.FontOptions;
                state.FontOptions = page.Options;
                try {
                    RenderSectionBlocks(document, section, pdf, options, state);
                    while (index + 1 < document.Sections.Count && !StartsNewPdfPage(document.Sections[index + 1].BreakKind)) {
                        index++;
                        RenderContinuousSectionBlocks(document, document.Sections[index], index, pdf, options, state, pdfOptions,
                            canContinueRunningStories, ref currentPageSection, ref currentPageOptions);
                    }
                } finally {
                    state.FontOptions = previousFontOptions;
                }
            });
        }
    }

    private static void RenderBlocks(RtfDocument document, IEnumerable<IRtfBlock> blocks, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        foreach (IRtfBlock block in blocks) {
            RenderBlock(document, block, pdf, options, state);
        }
    }

    private static bool StartsNewPdfPage(RtfSectionBreakKind breakKind) {
        switch (breakKind) {
            case RtfSectionBreakKind.Continuous:
                return false;
            default:
                return true;
        }
    }

    private static void RenderBlock(RtfDocument document, IRtfBlock block, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        options.CancellationToken.ThrowIfCancellationRequested();
        switch (block) {
            case RtfParagraph paragraph:
                state.PreviousBodyParagraph = paragraph;
                RenderParagraph(document, paragraph, pdf, options, state);
                break;
            case RtfTable table when options.IncludeTables:
                RenderTable(document, table, pdf, options, state);
                break;
            case RtfTable:
                AddConversionWarning(
                    options,
                    "TableSkipped",
                    "Table",
                    "An RTF table was skipped because IncludeTables is false.");
                break;
            case RtfImage image:
                RenderImage(image, pdf, options);
                break;
            case RtfObject rtfObject:
                AddConversionWarning(options, "ObjectFlattened", "Object", "RTF object was flattened to its visible text result.", RtfConversionAction.Flattened);
                RenderPlainTextBlock(rtfObject.ToPlainText(), pdf);
                break;
            case RtfShape shape:
                AddConversionWarning(options, "ShapeFlattened", "Shape", "RTF shape was flattened to its visible text result.", RtfConversionAction.Flattened);
                RenderPlainTextBlock(shape.ToPlainText(), pdf);
                break;
        }
    }

    private static void AddConversionWarning(RtfToPdfOptions options, string code, string source, string message, IReadOnlyDictionary<string, string>? details = null) {
        AddConversionWarning(options, code, source, message, RtfConversionAction.Omitted, details);
    }

    private static void AddConversionWarning(RtfToPdfOptions options, string code, string source, string message, RtfConversionAction action, IReadOnlyDictionary<string, string>? details = null) {
        var warningDetails = new Dictionary<string, string>();
        if (details != null) {
            foreach (KeyValuePair<string, string> detail in details) {
                warningDetails[detail.Key] = detail.Value;
            }
        }
        warningDetails["RtfAction"] = action.ToString();
        options.Report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Rtf.Pdf",
            code,
            source,
            message,
            details: warningDetails));
    }

    private static void RenderPlainTextBlock(string text, PdfCore.PdfDocument pdf) {
        if (!string.IsNullOrEmpty(text)) {
            pdf.Paragraph(paragraph => paragraph.Text(text));
        }
    }
}
