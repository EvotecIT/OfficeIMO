using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderContinuousSectionBlocks(RtfDocument document, RtfSection section, int sectionIndex,
        PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state, PdfCore.PdfOptions defaults,
        bool canContinueRunningStories, ref RtfSection currentPageSection, ref PdfCore.PdfOptions currentPageOptions) {
        if (canContinueRunningStories && TryPrepareContinuousRunningStoryOptions(document, section,
            currentPageSection, currentPageOptions, defaults, options, out PdfCore.PdfOptions? nextPageOptions)) {
            var page = new PdfCore.PdfPageBuilder(pdf, nextPageOptions!);
            ApplySectionHeaderFooters(document, section, page, pdf, options, state);
            // A source section can begin partway through a physical page. The
            // shared page owner retains that page's paint/frame and selects
            // this detached presentation only when a real later page starts.
            pdf.BeginRunningContentSection(sectionIndex, nextPageOptions!);
            currentPageOptions = nextPageOptions!;
            currentPageSection = section;
        } else {
            ReportContinuousSectionSettings(document, section, currentPageSection, currentPageOptions, defaults, options);
        }
        RenderSectionBlocks(document, section, pdf, options, state);
    }

    // Keep this admission separate from the established continuous-column
    // route. Notes can cross earlier boundaries, so inspect body ownership
    // once for the document rather than only the adjacent section pair.
    private static bool SupportsContinuousRunningStoryDocument(RtfDocument document) =>
        document.Settings.FacingPages != true && document.Settings.MirrorMargins != true &&
        document.Settings.Direction != RtfTextDirection.RightToLeft &&
        document.Sections.All(section =>
            (section.ColumnCount ?? Math.Max(1, section.Columns.Count)) == 1 &&
            section.Columns.Count == 0 && !section.ColumnSeparator &&
            section.Direction != RtfTextDirection.RightToLeft &&
            section.VerticalAlignment is null or RtfSectionVerticalAlignment.Top &&
            !section.LineNumbering.HasAnyValue &&
            !document.GetPageSetup(section).DifferentFirstPageHeaderFooter) &&
        document.Blocks.All(block => block is RtfParagraph paragraph &&
            !document.GetParagraphFormatting(paragraph).Frame.HasAnyValue &&
            paragraph.Inlines.All(inline => inline is RtfRun { Note: null } ||
                inline is RtfBreak { Kind: not RtfBreakKind.Column }));

    private static bool TryPrepareContinuousRunningStoryOptions(RtfDocument document, RtfSection section,
        RtfSection previousSection, PdfCore.PdfOptions previousOptions, PdfCore.PdfOptions defaults,
        RtfToPdfOptions options, out PdfCore.PdfOptions? nextPageOptions) {
        nextPageOptions = null;
        if (!options.IncludeHeaderFooters ||
            HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter: false) ||
            HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter: true) ||
            options.PdfOptions?.DifferentFirstPageHeaderFooter == true ||
            options.PdfOptions?.DifferentOddAndEvenPagesHeaderFooter == true ||
            section.PageSetup.PageNumberRestart == true ||
            (section.PageSetup.PageNumberStart.HasValue && section.PageSetup.PageNumberRestart != false)) return false;

        RtfPageSetup previousSetup = document.GetPageSetup(previousSection);
        RtfPageSetup requestedSetup = document.GetPageSetup(section);
        if (GetRunningStoryDistance(previousSetup.HeaderDistanceTwips) != GetRunningStoryDistance(requestedSetup.HeaderDistanceTwips) ||
            GetRunningStoryDistance(previousSetup.FooterDistanceTwips) != GetRunningStoryDistance(requestedSetup.FooterDistanceTwips)) return false;

        IReadOnlyList<RtfHeaderFooter> previousStories = document.GetEffectiveHeaderFooters(previousSection);
        IReadOnlyList<RtfHeaderFooter> requestedStories = document.GetEffectiveHeaderFooters(section);
        RtfHeaderFooter? previousHeader = GetDefaultHeaderFooter(previousStories, isFooter: false);
        RtfHeaderFooter? requestedHeader = GetDefaultHeaderFooter(requestedStories, isFooter: false);
        RtfHeaderFooter? previousFooter = GetDefaultHeaderFooter(previousStories, isFooter: true);
        if (previousHeader == null || requestedHeader == null || ReferenceEquals(previousHeader, requestedHeader) ||
            !ReferenceEquals(previousFooter, GetDefaultHeaderFooter(requestedStories, isFooter: true)) ||
            !IsPlainContinuousRunningStory(document, previousHeader) ||
            !IsPlainContinuousRunningStory(document, requestedHeader) ||
            !IsPlainContinuousRunningStory(document, previousFooter)) return false;

        // Reuse normal section setup on detached options. This neither adds
        // a PageBlock nor switches the current physical options or cursor.
        PdfCore.PdfOptions requested = defaults.Clone();
        ApplyPageSetup(document, section, requested, defaults);
        if (requested.PageWidth != previousOptions.PageWidth || requested.PageHeight != previousOptions.PageHeight ||
            requested.MarginLeft != previousOptions.MarginLeft || requested.MarginRight != previousOptions.MarginRight ||
            requested.MarginTop != previousOptions.MarginTop || requested.MarginBottom != previousOptions.MarginBottom ||
            requested.PageNumberStyle != previousOptions.PageNumberStyle ||
            requested.PageBorder != null || previousOptions.PageBorder != null) return false;

        // Both font collection and actual capture use this source admission.
        // Only the rendering caller captures rich blocks and diagnostics.
        nextPageOptions = requested;
        return true;
    }

    private static bool IsPlainContinuousRunningStory(RtfDocument document, RtfHeaderFooter? story) =>
        story == null || story.Paragraphs.All(paragraph =>
            !document.GetParagraphFormatting(paragraph).Frame.HasAnyValue &&
            !RtfPdfMapping.HasPageBreakBefore(document, paragraph) &&
            paragraph.Inlines.All(inline => inline is RtfRun { Note: null } ||
                inline is RtfBreak { Kind: RtfBreakKind.Line or RtfBreakKind.SoftLine }));
}
