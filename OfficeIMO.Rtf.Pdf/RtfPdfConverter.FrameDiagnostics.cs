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
        Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.RightHeader)
            ?? declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.Header));
        Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.RightFooter)
            ?? declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.Footer));
        if (setup.DifferentFirstPageHeaderFooter) {
            Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.FirstHeader));
            Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.FirstFooter));
        }
        if (document.Settings.FacingPages == true) {
            Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.LeftHeader));
            Record(declarations.LastOrDefault(story => story.Kind == RtfHeaderFooterKind.LeftFooter));
        }
        void Record(RtfHeaderFooter? story) {
            if (story == null) return;
            foreach (RtfParagraph paragraph in story.Paragraphs)
            {
                RecordParagraphFrameDiagnostic(document, paragraph, options, "HeaderFooter/Paragraph/Frame");
                RecordParagraphBarTabDiagnostic(document, paragraph, options, omitted: true);
            }
        }
    }
}
