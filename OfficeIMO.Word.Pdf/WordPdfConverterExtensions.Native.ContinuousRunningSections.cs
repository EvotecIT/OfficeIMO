using DocumentFormat.OpenXml;
using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Existing equivalent-story grouping is independent of this bounded
    // profile for transitions to a different default header.
    private static bool SupportsNativeContinuousRunningStoryDocument(
        WordDocument document, IReadOnlyList<WordSection> sections) {
        W.Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
        if (body == null || HasNativeContinuousStoryComplexContent(body)) return false;
        return sections.All(section =>
            (section.ColumnCount ?? 1) == 1 &&
            !section.DifferentFirstPage && !section.DocumentOddEvenSettingEnabled &&
            section._sectionProperties.GetFirstChild<W.PageNumberType>() == null &&
            section._sectionProperties.GetFirstChild<W.PageBorders>() == null);
    }

    // A note can continue from an earlier section. Check the source body once
    // rather than inspecting only the immediately adjacent section pair.
    private static bool HasNativeContinuousStoryComplexContent(OpenXmlElement element) =>
        element.Descendants().Any(descendant => descendant is
            W.Table or W.FootnoteReference or W.EndnoteReference or W.Drawing or
            W.Picture or W.TextBoxContent or W.EmbeddedObject or AlternateContent);

    private static bool CanContinueNativeDefaultHeaderStory(
        WordSection previous, WordSection current, WordPdfExportOptions? options) {
        if (HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter: false) ||
            HasNativeRunningStoryOverrides(options?.PdfOptions, isFooter: true) ||
            options?.PdfOptions?.DifferentFirstPageHeaderFooter == true ||
            options?.PdfOptions?.DifferentOddAndEvenPagesHeaderFooter == true ||
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
        return ReferenceEquals(previous.ResolveDefaultFooter()?._footer, current.ResolveDefaultFooter()?._footer) &&
            NativeDoubleEquals(previous.Margins.HeaderDistance, current.Margins.HeaderDistance) &&
            NativeDoubleEquals(previous.Margins.FooterDistance, current.Margins.FooterDistance);
    }

    private static bool IsNativePlainContinuousRunningStory(WordHeaderFooter? story) =>
        story == null ||
        (story.ChildElements.All(element => element is W.Paragraph) &&
         story.ChildElements.All(element =>
             !HasNativeContinuousStoryComplexContent(element) &&
             !element.Descendants<W.SimpleField>().Any() &&
             !element.Descendants<W.FieldChar>().Any() &&
             !element.Descendants<W.Break>().Any()));

    private static bool NativeSectionGroupChangesRunningStories(
        IReadOnlyList<WordSection> sections, int startIndex, int endIndex) {
        for (int index = startIndex + 1; index < endIndex; index++) {
            if (!NativeSectionHeaderFooterReferencesEquivalent(sections[index - 1], sections[index])) return true;
        }
        return false;
    }

    /// <summary>Reports an actual separate-page fallback for a continuous running-story change.</summary>
    private static void ReportNativeContinuousStoryPageFallback(WordSection previous, WordSection current,
        int sectionIndex, WordPdfExportOptions? options) {
        if (options == null || current.BreakType != WordSectionBreakType.Continuous ||
            !NativeContinuousRunningStorySelectionChanged(previous, current)) return;
        AddNativeExportWarning(options, "native-continuous-section-pagination-approximated",
            "continuous section " + (sectionIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
            "Changed header/footer stories start on a new physical page for this continuous-section layout profile. " +
            "Native continuous pagination is approximated.");
    }

    // Story identity and selection are distinct from edge distances and paper
    // geometry. A geometry-only refusal does not establish a story fallback.
    private static bool NativeContinuousRunningStorySelectionChanged(WordSection previous, WordSection current) =>
        previous.DifferentFirstPage != current.DifferentFirstPage ||
        previous.DocumentOddEvenSettingEnabled != current.DocumentOddEvenSettingEnabled ||
        !ReferenceEquals(previous.ResolveDefaultHeader()?._header, current.ResolveDefaultHeader()?._header) ||
        !ReferenceEquals(previous.ResolveDefaultFooter()?._footer, current.ResolveDefaultFooter()?._footer) ||
        (previous.DifferentFirstPage &&
            (!ReferenceEquals(previous.ResolveFirstHeader()?._header, current.ResolveFirstHeader()?._header) ||
             !ReferenceEquals(previous.ResolveFirstFooter()?._footer, current.ResolveFirstFooter()?._footer))) ||
        (previous.DocumentOddEvenSettingEnabled &&
            (!ReferenceEquals(previous.ResolveEvenHeader()?._header, current.ResolveEvenHeader()?._header) ||
             !ReferenceEquals(previous.ResolveEvenFooter()?._footer, current.ResolveEvenFooter()?._footer)));

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
