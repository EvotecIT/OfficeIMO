using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void ReportContinuousSectionSettings(RtfDocument document, RtfSection section, PdfCore.PdfOptions current, PdfCore.PdfOptions defaults, RtfToPdfOptions options) {
        PdfCore.PdfOptions requested = defaults.Clone();
        RtfPageSetup setup = document.GetPageSetup(section);
        ApplyPageSetup(document, setup, requested);
        // Do not emit the document-wide skipped-story diagnostic a second time.
        if (options.IncludeHeaderFooters) {
            ApplyHeaderFooters(document, setup, requested, options, document.GetEffectiveHeaderFooters(section));
        }
        var differences = new List<string>();
        if (requested.PageWidth != current.PageWidth || requested.PageHeight != current.PageHeight) differences.Add("PaperSize");
        if (requested.MarginLeft != current.MarginLeft || requested.MarginRight != current.MarginRight ||
            requested.MarginTop != current.MarginTop || requested.MarginBottom != current.MarginBottom) differences.Add("Margins");
        if (requested.HeaderFormat != current.HeaderFormat || requested.FirstPageHeaderFormat != current.FirstPageHeaderFormat ||
            requested.EvenPageHeaderFormat != current.EvenPageHeaderFormat || requested.ShowHeader != current.ShowHeader) differences.Add("Headers");
        if (requested.FooterFormat != current.FooterFormat || requested.FirstPageFooterFormat != current.FirstPageFooterFormat ||
            requested.EvenPageFooterFormat != current.EvenPageFooterFormat || requested.ShowPageNumbers != current.ShowPageNumbers) differences.Add("Footers");
        if (requested.DifferentFirstPageHeaderFooter != current.DifferentFirstPageHeaderFooter ||
            requested.DifferentOddAndEvenPagesHeaderFooter != current.DifferentOddAndEvenPagesHeaderFooter) differences.Add("StorySelection");
        if (requested.PageNumberStyle != current.PageNumberStyle || section.PageSetup.PageNumberRestart == true ||
            (section.PageSetup.PageNumberStart.HasValue && section.PageSetup.PageNumberRestart != false)) differences.Add("PageNumbering");
        PdfCore.PdfPageBorder? requestedBorder = requested.PageBorder;
        PdfCore.PdfPageBorder? currentBorder = current.PageBorder;
        if (requestedBorder == null ? currentBorder != null : currentBorder == null ||
            !requestedBorder.Color.Equals(currentBorder.Color) || requestedBorder.Width != currentBorder.Width ||
            requestedBorder.Inset != currentBorder.Inset || requestedBorder.Opacity != currentBorder.Opacity ||
            requestedBorder.DashStyle != currentBorder.DashStyle) differences.Add("PageBorders");
        if (differences.Count == 0) return;
        AddConversionWarning(options, "ContinuousSectionPageSettingsFlattened", "Section/PageSetup",
            "Continuous RTF section page settings retain the current PDF page settings; column layout still applies.",
            RtfConversionAction.Flattened, new Dictionary<string, string> { ["Properties"] = string.Join(",", differences) });
    }

    private static void RenderSectionBlocks(RtfDocument document, RtfSection section, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        int count = section.ColumnCount ?? Math.Max(1, section.Columns.Count);
        if (count <= 1) {
            RenderBlocks(document, section.Blocks, pdf, options, state);
            return;
        }
        if (section.Columns.Count > 0 || count > 12 || section.ColumnSpaceTwips < 0 ||
            (section.Direction ?? document.Settings.Direction) == RtfTextDirection.RightToLeft) {
            AddConversionWarning(options, "SectionColumnsFlattened", "Section/Columns",
                "Unequal, right-to-left, or unsupported RTF section columns were flattened to ordinary flow.", RtfConversionAction.Flattened);
            RenderBlocks(document, section.Blocks, pdf, options, state);
            return;
        }
        var columns = new PdfCore.PdfMultiColumnOptions {
            ColumnCount = count,
            Gap = RtfPdfMapping.TwipsToPoints(section.ColumnSpaceTwips ?? 720),
            SeparatorColor = section.ColumnSeparator ? new PdfCore.PdfColor(0, 0, 0) : null,
            SeparatorWidth = section.ColumnSeparator ? 0.5 : 0
        };
        pdf.Columns(_ => {
            state.InColumns = true;
            try {
                RenderBlocks(document, section.Blocks, pdf, options, state);
            } finally {
                state.InColumns = false;
            }
        }, columns);
    }
}
