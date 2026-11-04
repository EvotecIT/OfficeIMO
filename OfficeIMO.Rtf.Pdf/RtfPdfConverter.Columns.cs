using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void ReportContinuousSectionSettings(RtfDocument document, RtfSection section, RtfSection currentPageSection,
        PdfCore.PdfCreateOptions current, PdfCore.PdfCreateOptions defaults, RtfPdfExportOptions options) {
        PdfCore.PdfCreateOptions requested = defaults.Clone();
        RtfPageSetup setup = document.GetPageSetup(section);
        ApplyPageSetup(document, setup, requested);
        // Do not emit the document-wide skipped-story diagnostic a second time.
        if (options.IncludeHeaderFooters) {
            ApplyHeaderFooters(document, setup, requested, options, document.GetEffectiveHeaderFooters(section));
        }
        var differences = new List<string>();
        if (requested.PageSetup.Size.Width != current.PageSetup.Size.Width || requested.PageSetup.Size.Height != current.PageSetup.Size.Height) differences.Add("PaperSize");
        if (requested.PageSetup.Margins.Left != current.PageSetup.Margins.Left || requested.PageSetup.Margins.Right != current.PageSetup.Margins.Right ||
            requested.PageSetup.Margins.Top != current.PageSetup.Margins.Top || requested.PageSetup.Margins.Bottom != current.PageSetup.Margins.Bottom) differences.Add("Margins");
        if (requested.PageSetup.Header.Format != current.PageSetup.Header.Format || requested.PageSetup.FirstPageHeader.Format != current.PageSetup.FirstPageHeader.Format ||
            requested.PageSetup.EvenPageHeader.Format != current.PageSetup.EvenPageHeader.Format || requested.PageSetup.Header.Show != current.PageSetup.Header.Show ||
            RunningStoriesDiffer(document, currentPageSection, section, options, isFooter: false)) differences.Add("Headers");
        if (requested.PageSetup.Footer.Format != current.PageSetup.Footer.Format || requested.PageSetup.FirstPageFooter.Format != current.PageSetup.FirstPageFooter.Format ||
            requested.PageSetup.EvenPageFooter.Format != current.PageSetup.EvenPageFooter.Format || requested.PageSetup.Footer.Show != current.PageSetup.Footer.Show ||
            RunningStoriesDiffer(document, currentPageSection, section, options, isFooter: true)) differences.Add("Footers");
        if (requested.PageSetup.DifferentFirstPage != current.PageSetup.DifferentFirstPage ||
            requested.PageSetup.DifferentOddAndEvenPages != current.PageSetup.DifferentOddAndEvenPages) differences.Add("StorySelection");
        if (requested.PageNumberStyle != current.PageNumberStyle || section.PageSetup.PageNumberRestart == true ||
            (section.PageSetup.PageNumberStart.HasValue && section.PageSetup.PageNumberRestart != false)) differences.Add("PageNumbering");
        PdfCore.PdfPageBorder? requestedBorder = requested.PageBorder;
        PdfCore.PdfPageBorder? currentBorder = current.PageBorder;
        if (requestedBorder == null ? currentBorder != null : currentBorder == null ||
            !requestedBorder.Color.Equals(currentBorder.Color) || requestedBorder.Width.ToPoints() != currentBorder.Width.ToPoints() ||
            requestedBorder.Inset.ToPoints() != currentBorder.Inset.ToPoints() || requestedBorder.Opacity != currentBorder.Opacity ||
            requestedBorder.DashStyle != currentBorder.DashStyle) differences.Add("PageBorders");
        if (differences.Count == 0) return;
        AddConversionWarning(options, "ContinuousSectionPageSettingsFlattened", "Section/PageSetup",
            "Continuous RTF section page settings retain the current PDF page settings; column layout still applies.",
            RtfConversionAction.Flattened, new Dictionary<string, string> { ["Properties"] = string.Join(",", differences) });
    }

    private static void RenderSectionBlocks(RtfDocument document, RtfSection section, PdfCore.PdfDocument pdf, RtfPdfExportOptions options, PdfRenderState state) {
        // A section terminator owns an implicit paragraph even when an
        // authored model has no explicit body blocks. Active body paragraph
        // controls cross section boundaries; the first uses normal parsed
        // defaults. Detached formatting never changes the source collection.
        IEnumerable<IRtfBlock> blocks = section.Blocks.Count == 0
            ? new[] { state.CreateImplicitBodyParagraph(document) }
            : section.Blocks;
        int count = section.ColumnCount ?? Math.Max(1, section.Columns.Count);
        if (count <= 1) {
            RenderBlocks(document, blocks, pdf, options, state);
            return;
        }
        if (section.Columns.Count > 0 || count > 12 || section.ColumnSpaceTwips < 0 ||
            (section.Direction ?? document.Settings.Direction) == RtfTextDirection.RightToLeft) {
            AddConversionWarning(options, "SectionColumnsFlattened", "Section/Columns",
                "Unequal, right-to-left, or unsupported RTF section columns were flattened to ordinary flow.", RtfConversionAction.Flattened);
            RenderBlocks(document, blocks, pdf, options, state);
            return;
        }
        var columns = new PdfCore.PdfMultiColumnOptions {
            ColumnCount = count,
            Gap = global::OfficeIMO.OfficeLength.Points(RtfPdfMapping.TwipsToPoints(section.ColumnSpaceTwips ?? 720)),
            SeparatorColor = section.ColumnSeparator ? global::OfficeIMO.OfficeColor.FromNormalizedRgb(0, 0, 0) : null,
            SeparatorWidth = global::OfficeIMO.OfficeLength.Points(section.ColumnSeparator ? 0.5 : 0)
        };
        pdf.Columns(_ => {
            state.InColumns = true;
            try {
                RenderBlocks(document, blocks, pdf, options, state);
            } finally {
                state.InColumns = false;
            }
        }, columns);
    }
}
