using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Existing equivalent-story grouping is independent of this bounded
    // profile for transitions to a different default header.
    private static bool SupportsNativeContinuousRunningStoryDocument(
        WordDocument document, IReadOnlyList<WordSection> sections) {
        OfficeXmlNode body = document.BodyNode();
        if (body.IsNone || HasNativeContinuousStoryComplexContent(body)) return false;
        return sections.All(section =>
            (section.ColumnCount ?? 1) == 1 &&
            !section.PageSetup.DifferentFirstPage && !section.DocumentOddEvenSettingEnabled &&
            section.Node().WChild("pgNumType").IsNone &&
            section.Node().WChild("pgBorders").IsNone);
    }

    // A note can continue from an earlier section. Check the source body once
    // rather than inspecting only the immediately adjacent section pair.
    private static bool HasNativeContinuousStoryComplexContent(OfficeXmlNode element) =>
        element.DescendantElements().Any(descendant =>
            descendant.IsW("tbl") || descendant.IsW("footnoteReference") || descendant.IsW("endnoteReference") ||
            descendant.IsW("drawing") || descendant.IsW("pict") || descendant.IsW("txbxContent") ||
            descendant.IsW("object") || descendant.Is(WordXml.Mc, "AlternateContent"));

    private static bool CanContinueNativeDefaultHeaderStory(
        WordSection previous, WordSection current, WordPdfExportOptions? options) {
        if (HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter: false) ||
            HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter: true) ||
            options?.PdfOptions?.PageSetup.DifferentFirstPage == true ||
            options?.PdfOptions?.PageSetup.DifferentOddAndEvenPages == true ||
            ShouldMirrorNativeMargins(previous, options) != ShouldMirrorNativeMargins(current, options)) return false;

        WordHeaderFooter? previousHeader = previous.ResolveDefaultHeader();
        WordHeaderFooter? currentHeader = current.ResolveDefaultHeader();
        if (previousHeader == null || currentHeader == null ||
            !UsesNativeRunningHeader(previous, options) || !UsesNativeRunningHeader(current, options) ||
            !IsNativePlainContinuousRunningStory(previousHeader) ||
            !IsNativePlainContinuousRunningStory(currentHeader) ||
            !IsNativePlainContinuousRunningStory(previous.ResolveDefaultFooter())) return false;

        // A default-header transition leaves the selected footer and both
        // running edges fixed. First/even variants and numbering stay outside
        // this source profile.
        return (previous.ResolveDefaultFooter()?.Node() ?? default) == (current.ResolveDefaultFooter()?.Node() ?? default) &&
            NativeDoubleEquals(checked((uint)(previous.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips())), checked((uint)(current.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips()))) &&
            NativeDoubleEquals(checked((uint)(previous.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips())), checked((uint)(current.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips())));
    }

    private static bool IsNativePlainContinuousRunningStory(WordHeaderFooter? story) =>
        story == null ||
        (story.Node().Elements().All(element => element.IsW("p")) &&
         story.Node().Elements().All(element =>
             !HasNativeContinuousStoryComplexContent(element) &&
             !element.HasDescendant(WordXml.W, "fldSimple") &&
             !element.HasDescendant(WordXml.W, "fldChar") &&
             !element.HasDescendant(WordXml.W, "br")));

    private static bool NativeSectionGroupChangesRunningStories(
        IReadOnlyList<WordSection> sections, int startIndex, int endIndex) {
        for (int index = startIndex + 1; index < endIndex; index++) {
            if (!NativeSectionHeaderFooterReferencesEquivalent(sections[index - 1], sections[index])) return true;
        }
        return false;
    }

    /// <summary>Captures the resolved stories for a later physical page without appending a page or evaluating running factories.</summary>
    private static PdfCore.PdfCreateOptions CreateNativeContinuousRunningOptions(
        PdfCore.PdfDocument pdf, WordSection section, WordPdfExportOptions? options,
        NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        NativeDocumentDefaults defaults) {
        PdfCore.PdfCreateOptions snapshot = pdf.Options.Clone();
        snapshot.ClearPageNumberStartOverride();
        snapshot.SetRunningContent(null, isHeader: true, PdfCore.PdfRunningContentVariant.Default);
        snapshot.SetRunningContent(null, isHeader: false, PdfCore.PdfRunningContentVariant.Default);
        snapshot.SetRunningContent(null, isHeader: true, PdfCore.PdfRunningContentVariant.FirstPage);
        snapshot.SetRunningContent(null, isHeader: false, PdfCore.PdfRunningContentVariant.FirstPage);
        snapshot.SetRunningContent(null, isHeader: true, PdfCore.PdfRunningContentVariant.EvenPages);
        snapshot.SetRunningContent(null, isHeader: false, PdfCore.PdfRunningContentVariant.EvenPages);
        var page = new PdfCore.PdfPageBuilder(pdf, snapshot);
        (double Header, double Footer) expansion = GetNativeHeaderFooterMarginExpansion(section, options, listMarkers);
        page.Size(GetNativePageSize(section, options));
        page.Margin(GetNativeMargins(section, options, expansion));
        page.MirrorMargins(ShouldMirrorNativeMargins(section, options));
        page.StartOnPageParity(null);
        // Numbering changes are not admitted here; retain continuing visible
        // numbering while the PDF owner measures the new physical frame.
        ConfigureNativeHeaderFooter(page, section, options, expansion.Header, expansion.Footer, fontMap, listMarkers);
        ConfigureNativeRunningHeaderFooter(page, section, options, fontMap, listMarkers, defaults);
        return snapshot;
    }
}
