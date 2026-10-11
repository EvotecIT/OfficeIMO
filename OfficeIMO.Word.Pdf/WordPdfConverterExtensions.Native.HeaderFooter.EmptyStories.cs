using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static WordParagraph CreateNativeImplicitRunningParagraph(WordHeaderFooter story) =>
        new(story.Document, new W.Paragraph());

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
