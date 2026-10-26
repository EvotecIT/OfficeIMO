using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static RtfHeaderFooter? GetHeaderFooter(IReadOnlyList<RtfHeaderFooter> declarations, RtfHeaderFooterKind kind) =>
        declarations.LastOrDefault(story => story.Kind == kind);

    private static RtfHeaderFooter? GetDefaultHeaderFooter(IReadOnlyList<RtfHeaderFooter> declarations, bool isFooter) =>
        GetHeaderFooter(declarations, isFooter ? RtfHeaderFooterKind.RightFooter : RtfHeaderFooterKind.RightHeader)
        ?? GetHeaderFooter(declarations, isFooter ? RtfHeaderFooterKind.Footer : RtfHeaderFooterKind.Header);

    private static IEnumerable<RtfHeaderFooter> GetSelectedHeaderFooters(RtfDocument document, RtfPageSetup setup,
        IReadOnlyList<RtfHeaderFooter> declarations, bool isFooter) {
        RtfHeaderFooter? story = GetDefaultHeaderFooter(declarations, isFooter);
        if (story != null) yield return story;
        story = GetHeaderFooter(declarations, isFooter ? RtfHeaderFooterKind.FirstFooter : RtfHeaderFooterKind.FirstHeader);
        if (setup.DifferentFirstPageHeaderFooter && story != null) yield return story;
        story = GetHeaderFooter(declarations, isFooter ? RtfHeaderFooterKind.LeftFooter : RtfHeaderFooterKind.LeftHeader);
        if (document.Settings.FacingPages == true && story != null) yield return story;
    }

    /// <summary>Retains the established plain-story presentation when callers explicitly configure that projection.</summary>
    private static bool HasRunningStoryPresentationOverrides(PdfCore.PdfOptions? options, bool isFooter) => options != null &&
        (isFooter
            ? options.FooterOffsetY != 18D || options.HasExplicitFooterFont || options.FooterFontFamily != null ||
                options.FooterFontSize != 9D || options.FooterTextColor.HasValue || options.FooterAlign != PdfCore.PdfAlign.Center
            : options.HeaderOffsetY != 18D || options.HasExplicitHeaderFont || options.HeaderFontFamily != null ||
                options.HeaderFontSize != 9D || options.HeaderTextColor.HasValue || options.HeaderAlign != PdfCore.PdfAlign.Center);

    /// <summary>Captures rich source stories once through the same document and paragraph/image owners used by ordinary flow.</summary>
    private static void ApplyHeaderFooterContent(RtfDocument document, RtfPageSetup setup, PdfCore.PdfDocument pdf,
        PdfCore.PdfOptions pageOptions, RtfToPdfOptions options, PdfRenderState state,
        IReadOnlyList<RtfHeaderFooter> declarations, Action<Action<PdfCore.PdfHeaderBuilder>> configureHeader,
        Action<Action<PdfCore.PdfFooterBuilder>> configureFooter) {
        if (!options.IncludeHeaderFooters || declarations.Count == 0) return;

        if (!HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter: false)) {
            double distance = GetRunningStoryDistance(setup.HeaderDistanceTwips);
            configureHeader(header => {
                RtfHeaderFooter? story = GetDefaultHeaderFooter(declarations, isFooter: false);
                if (story != null) header.Content(CreateRunningStory(story), distance);
                story = GetHeaderFooter(declarations, RtfHeaderFooterKind.FirstHeader);
                if (setup.DifferentFirstPageHeaderFooter && story != null) header.FirstPageContent(CreateRunningStory(story), distance);
                story = GetHeaderFooter(declarations, RtfHeaderFooterKind.LeftHeader);
                if (document.Settings.FacingPages == true && story != null) header.EvenPagesContent(CreateRunningStory(story), distance);
            });
        } else {
            RecordPlainRunningImageDiagnostics(isFooter: false);
        }
        if (!HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter: true)) {
            double distance = GetRunningStoryDistance(setup.FooterDistanceTwips);
            configureFooter(footer => {
                RtfHeaderFooter? story = GetDefaultHeaderFooter(declarations, isFooter: true);
                if (story != null) footer.Content(CreateRunningStory(story), distance);
                story = GetHeaderFooter(declarations, RtfHeaderFooterKind.FirstFooter);
                if (setup.DifferentFirstPageHeaderFooter && story != null) footer.FirstPageContent(CreateRunningStory(story), distance);
                story = GetHeaderFooter(declarations, RtfHeaderFooterKind.LeftFooter);
                if (document.Settings.FacingPages == true && story != null) footer.EvenPagesContent(CreateRunningStory(story), distance);
            });
        } else {
            RecordPlainRunningImageDiagnostics(isFooter: true);
        }
        // Content builders enable their variant. Source flags remain the
        // authority, including explicitly empty and absent selected stories.
        pageOptions.DifferentFirstPageHeaderFooter = setup.DifferentFirstPageHeaderFooter;
        pageOptions.DifferentOddAndEvenPagesHeaderFooter = document.Settings.FacingPages == true;
        pageOptions.UsePageNumberParityForHeaderFooter = true;

        Action<PdfCore.PdfContentBuilder> CreateRunningStory(RtfHeaderFooter story) {
            PdfRenderState runningState = state.CreateRunningStoryState(document, pageOptions);
            return _ => {
                foreach (RtfParagraph paragraph in story.Paragraphs) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    RecordRunningParagraphDiagnostics(document, paragraph, options);
                    var runs = new List<PdfCore.PdfTextRun>();
                    AppendParagraphRuns(document, paragraph, runs, options, runningState, collectNotes: false,
                        context: PdfInlineProjectionContext.RunningStory);
                    EnsureParagraphLineBoxStrut(document, paragraph, runs, runningState);
                    pdf.Paragraph(builder => builder.Runs(runs),
                        RtfPdfMapping.ToPdfAlign(document.GetParagraphFormatting(paragraph).Alignment),
                        style: RtfPdfMapping.ToPdfParagraphStyle(document, paragraph, runs, pageOptions));
                }
            };
        }

        void RecordPlainRunningImageDiagnostics(bool isFooter) {
            foreach (RtfHeaderFooter story in GetSelectedHeaderFooters(document, setup, declarations, isFooter))
                foreach (RtfParagraph paragraph in story.Paragraphs) RecordImages(paragraph);

            void RecordImages(RtfParagraph paragraph) {
                if (!options.DiagnosedRunningParagraphs.Add(paragraph)) return;
                foreach (RtfImage _ in paragraph.Inlines.OfType<RtfImage>()) {
                    AddConversionWarning(options, "ImageSkipped", "HeaderFooter/Paragraph/Image",
                        "An RTF running-story image was skipped because caller-configured plain header/footer presentation is selected.");
                }
                foreach (RtfField field in paragraph.Inlines.OfType<RtfField>()) RecordImages(field.Result);
            }
        }
    }

    // The native RTF default is half an inch. Negative malformed source
    // distances retain the page edge rather than violating bounded flow.
    private static double GetRunningStoryDistance(int? twips) => Math.Max(0D, (twips ?? 720) / 20D);

    /// <summary>Reports unsupported running controls before static capture; no warning depends on pagination replay.</summary>
    private static void RecordRunningParagraphDiagnostics(RtfDocument document, RtfParagraph paragraph, RtfToPdfOptions options) {
        if (!options.DiagnosedRunningParagraphs.Add(paragraph)) return;
        RecordParagraphFrameDiagnostic(document, paragraph, options, "HeaderFooter/Paragraph/Frame");
        RecordParagraphBarTabDiagnostic(document, paragraph, options);
        if (RtfPdfMapping.HasPageBreakBefore(document, paragraph) || paragraph.Inlines.OfType<RtfBreak>()
            .Any(item => item.Kind is RtfBreakKind.Page or RtfBreakKind.SoftPage or RtfBreakKind.Column)) {
            AddConversionWarning(options, "HeaderFooterBreakFlattened", "HeaderFooter/Paragraph/Break",
                "RTF running-story page-break-before is ignored and inline page/column breaks use line breaks; bounded content cannot create pages or columns.",
                RtfConversionAction.Flattened);
        }
        if (paragraph.Inlines.Any(inline => inline is RtfRun { Note: not null } || inline is RtfGeneratedText { Note: not null })) {
            AddConversionWarning(options, "HeaderFooterNoteBodyOmitted", "HeaderFooter/Paragraph/Note",
                "RTF running-story note markers retain their visible text; their note bodies are not appended to document flow.");
        }
        if (paragraph.Inlines.OfType<RtfBookmarkMarker>().Any()) {
            AddConversionWarning(options, "HeaderFooterBookmarkOmitted", "HeaderFooter/Paragraph/Bookmark",
                "RTF running-story bookmark destinations are not repeated into document flow.");
        }
        if (paragraph.Inlines.OfType<RtfGeneratedText>().Any(item => item.Note == null)) {
            AddConversionWarning(options, "HeaderFooterGeneratedTextFallback", "HeaderFooter/Paragraph/GeneratedText",
                "RTF running-story generated text uses its cached fallback; values are not evaluated for each PDF page.",
                RtfConversionAction.Flattened);
        }
        foreach (RtfField field in paragraph.Inlines.OfType<RtfField>()) {
            if (!string.Equals(field.FieldCode.Keyword, "HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                AddConversionWarning(options, "HeaderFooterFieldResultUsed", "HeaderFooter/Paragraph/Field",
                    "RTF running-story fields use their rich cached result; field instructions are not evaluated for each PDF page.",
                    RtfConversionAction.Flattened);
            }
            RecordRunningParagraphDiagnostics(document, field.Result, options);
        }
    }

    /// <summary>Detects replaced rich destinations and edge geometry even when their flattened text is unchanged.</summary>
    private static bool RunningStoriesDiffer(RtfDocument document, RtfSection currentSection, RtfSection requestedSection,
        RtfToPdfOptions options, bool isFooter) {
        if (!options.IncludeHeaderFooters || HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter)) return false;
        IReadOnlyList<RtfHeaderFooter> current = document.GetEffectiveHeaderFooters(currentSection);
        IReadOnlyList<RtfHeaderFooter> requested = document.GetEffectiveHeaderFooters(requestedSection);
        RtfPageSetup currentSetup = document.GetPageSetup(currentSection);
        RtfPageSetup requestedSetup = document.GetPageSetup(requestedSection);
        if (!ReferenceEquals(GetDefaultHeaderFooter(current, isFooter), GetDefaultHeaderFooter(requested, isFooter))) return true;
        if (requestedSetup.DifferentFirstPageHeaderFooter && !ReferenceEquals(
            GetHeaderFooter(current, isFooter ? RtfHeaderFooterKind.FirstFooter : RtfHeaderFooterKind.FirstHeader),
            GetHeaderFooter(requested, isFooter ? RtfHeaderFooterKind.FirstFooter : RtfHeaderFooterKind.FirstHeader))) return true;
        if (document.Settings.FacingPages == true && !ReferenceEquals(
            GetHeaderFooter(current, isFooter ? RtfHeaderFooterKind.LeftFooter : RtfHeaderFooterKind.LeftHeader),
            GetHeaderFooter(requested, isFooter ? RtfHeaderFooterKind.LeftFooter : RtfHeaderFooterKind.LeftHeader))) return true;
        return (GetSelectedHeaderFooters(document, currentSetup, current, isFooter).Any() ||
            GetSelectedHeaderFooters(document, requestedSetup, requested, isFooter).Any()) &&
            GetRunningStoryDistance(isFooter ? currentSetup.FooterDistanceTwips : currentSetup.HeaderDistanceTwips)
            != GetRunningStoryDistance(isFooter ? requestedSetup.FooterDistanceTwips : requestedSetup.HeaderDistanceTwips);
    }
}
