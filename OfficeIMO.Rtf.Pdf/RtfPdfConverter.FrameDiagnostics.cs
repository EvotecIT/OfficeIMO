namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    /// <summary>Records lost effective frame layout only for paragraphs projected by this operation.</summary>
    private static void RecordParagraphFrameDiagnostic(RtfDocument document, RtfParagraph paragraph,
        RtfToPdfOptions options, string source) {
        if (!options.DiagnosedFrameParagraphs.Add(paragraph) ||
            !document.GetParagraphFormatting(paragraph).Frame.HasAnyValue) return;
        AddConversionWarning(options, "ParagraphFrameFlattened", source,
            "RTF paragraph frame positioning, dimensions and wrapping are not supported by PDF conversion; its visible content is placed in ordinary flow.",
            RtfConversionAction.Flattened);
    }

    /// <summary>Uses the same effective story selection as running-text projection.</summary>
    private static void RecordHeaderFooterFrameDiagnostics(RtfDocument document, RtfPageSetup setup,
        IReadOnlyList<RtfHeaderFooter> declarations, RtfToPdfOptions options) {
        if (!options.IncludeHeaderFooters) return;
        Record(GetDefaultHeaderFooter(declarations, isFooter: false), isFooter: false);
        Record(GetDefaultHeaderFooter(declarations, isFooter: true), isFooter: true);
        if (setup.DifferentFirstPageHeaderFooter) {
            Record(GetHeaderFooter(declarations, RtfHeaderFooterKind.FirstHeader), isFooter: false);
            Record(GetHeaderFooter(declarations, RtfHeaderFooterKind.FirstFooter), isFooter: true);
        }
        if (document.Settings.FacingPages == true) {
            Record(GetHeaderFooter(declarations, RtfHeaderFooterKind.LeftHeader), isFooter: false);
            Record(GetHeaderFooter(declarations, RtfHeaderFooterKind.LeftFooter), isFooter: true);
        }
        void Record(RtfHeaderFooter? story, bool isFooter) {
            if (story == null) return;
            foreach (RtfParagraph paragraph in story.Paragraphs)
            {
                RecordParagraphFrameDiagnostic(document, paragraph, options, "HeaderFooter/Paragraph/Frame");
                RecordParagraphBarTabDiagnostic(document, paragraph, options,
                    omittedSource: HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter) ? "HeaderFooter/Paragraph/Tabs" : null);
            }
        }
    }
}
