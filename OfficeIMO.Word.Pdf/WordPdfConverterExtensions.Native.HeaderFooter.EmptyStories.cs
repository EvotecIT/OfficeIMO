using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static WordParagraph CreateNativeImplicitRunningParagraph(WordHeaderFooter story) =>
        new(story.Document, new W.Paragraph());

    private static void ConfigureNativeImplicitRunningStories(PdfCore.PdfPageBuilder page, WordSection section,
        WordToPdfOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        NativeDocumentDefaults defaults, bool isFooter) {
        WordHeaderFooter? defaultStory = isFooter ? section.ResolveDefaultFooter() : section.ResolveDefaultHeader();
        WordHeaderFooter? firstStory = section.DifferentFirstPage
            ? isFooter ? section.ResolveFirstFooter() : section.ResolveFirstHeader() : null;
        WordHeaderFooter? evenStory = section.DocumentOddEvenSettingEnabled
            ? isFooter ? section.ResolveEvenFooter() : section.ResolveEvenHeader() : null;
        WordHeaderFooter?[] stories = { defaultStory, firstStory, evenStory };
        double distance = (isFooter ? section.Margins.FooterDistance : section.Margins.HeaderDistance) / 20D;
        for (int index = 0; index < stories.Length; index++) {
            WordHeaderFooter? story = stories[index];
            if (story == null || story.ChildElements.Any()) continue;
            // An implicit source paragraph still owns its blank line geometry
            // when visible story text uses the caller-configured projection.
            // Configure only this variant; its siblings retain their overrides.
            var content = CreateNativeRunningStory(story, section, options, fontMap, listMarkers, defaults, isFooter,
                includeGeneratedPageNumber: false);
            if (isFooter) page.Footer(footer => {
                if (index == 1) footer.FirstPageContent(content, distance);
                else if (index == 2) footer.EvenPagesContent(content, distance);
                else footer.Content(content, distance);
            });
            else page.Header(header => {
                if (index == 1) header.FirstPageContent(content, distance);
                else if (index == 2) header.EvenPagesContent(content, distance);
                else header.Content(content, distance);
            });
            if (isFooter) RetainNativeTextFreeFooterProjection(page.Options, story, listMarkers, fontMap,
                index == 1 ? PdfCore.PdfRunningContentVariant.FirstPage
                    : index == 2 ? PdfCore.PdfRunningContentVariant.EvenPages : PdfCore.PdfRunningContentVariant.Default);
        }
        ReserveNativeEmptyRunningContentDistances(page.Options, !isFooter, defaultStory, firstStory, evenStory);
        page.Options.AllowRunningContentTextFallback(!isFooter);
    }

    private static void RetainNativeTextFreeFooterProjection(PdfCore.PdfOptions options, WordHeaderFooter? story,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap fontMap,
        PdfCore.PdfRunningContentVariant variant) {
        // Keep generated default numbers at their physical margin. The native
        // definition still owns table/object drawing and its body exclusion;
        // retaining text does not enable a second flat image/shape projection.
        if (story != null && GetNativeHeaderFooterText(story, listMarkers, fontMap) == null)
            options.RetainRunningContentTextProjection(isHeader: false, variant);
    }

    private static void RenderNativeImplicitRunningParagraph(INativePdfFlow flow, WordHeaderFooter story,
        NativeDocumentDefaults defaults, NativeFontMap fontMap) {
        // Word repairs an empty referenced part with a default-paragraph line.
        // It inherits typography and spacing before, but omits spacing after.
        // Keep the synthetic mark detached so conversion leaves the part unchanged.
        WordParagraph paragraph = CreateNativeImplicitRunningParagraph(story);
        if (ShouldRenderNativeEmptyParagraphLineBox(paragraph)) {
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
            flow.Spacer(style.SpacingBefore + MeasureNativeEmptyParagraphLineHeight(paragraph, defaults, fontMap));
        }
    }

    private static void ReserveNativeEmptyRunningContentDistances(PdfCore.PdfOptions options, bool isHeader,
        WordHeaderFooter? defaultStory, WordHeaderFooter? firstStory, WordHeaderFooter? evenStory) {
        if (defaultStory != null)
            options.ReserveEmptyRunningContentDistance(isHeader, PdfCore.PdfRunningContentVariant.Default);
        if (firstStory != null)
            options.ReserveEmptyRunningContentDistance(isHeader, PdfCore.PdfRunningContentVariant.FirstPage);
        if (evenStory != null)
            options.ReserveEmptyRunningContentDistance(isHeader, PdfCore.PdfRunningContentVariant.EvenPages);
    }
}
