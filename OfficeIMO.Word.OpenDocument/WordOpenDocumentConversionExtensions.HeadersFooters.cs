using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static IEnumerable<(OdtHeaderFooter? Story, WordHeaderFooterType Kind, bool IsHeader)> EnumerateOdtHeaderFooterVariants(OdtPageLayout layout) {
        yield return (layout.HasHeader ? layout.Header : null, WordHeaderFooterType.Default, true);
        yield return (layout.HasFooter ? layout.Footer : null, WordHeaderFooterType.Default, false);
        yield return (layout.FirstHeader, WordHeaderFooterType.First, true);
        yield return (layout.FirstFooter, WordHeaderFooterType.First, false);
        yield return (layout.LeftHeader, WordHeaderFooterType.Even, true);
        yield return (layout.LeftFooter, WordHeaderFooterType.Even, false);
    }

    private static IEnumerable<OdtHeaderFooter> EnumerateOdtHeaderFooters(OdtPageLayout layout) {
        foreach ((OdtHeaderFooter? story, _, _) in EnumerateOdtHeaderFooterVariants(layout))
            if (story != null) yield return story;
    }

    private static OdtHeaderFooter? ResolveOdtHeaderFooterFallback(OdtPageLayout layout,
        WordHeaderFooterType kind, bool isHeader) {
        bool alternateApplies = kind == WordHeaderFooterType.First
            ? layout.FirstHeader != null || layout.FirstFooter != null
            : kind == WordHeaderFooterType.Even && (layout.LeftHeader != null || layout.LeftFooter != null);
        if (!alternateApplies) return null;
        return isHeader ? (layout.HasHeader ? layout.Header : null) : (layout.HasFooter ? layout.Footer : null);
    }

    private static void CopyOdtHeaderFooter(OdtHeaderFooter source, WordHeaderFooter target,
        WordOpenDocumentConversionOptions options, CultureInfo textCaseCulture,
        ref int hyperlinks, ref int externalHyperlinks, ref int images, ref int bookmarks,
        ref int approximatedRuns, ref int approximatedBookmarkRanges, ref int unsupportedMeasurements,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies,
        ref int mappedFields, ref int unsupportedFields,
        HashSet<XElement> handledUnsupportedFieldElements, NoteMappingStats notes) {
        foreach (OdtParagraph paragraph in source.Paragraphs) {
            WordParagraph converted = target.AddParagraph();
            if (paragraph.HeadingLevel.HasValue)
                converted.Style = HeadingStyle(paragraph.HeadingLevel.Value);
            CopyParagraph(paragraph, converted, options, textCaseCulture, ref hyperlinks, ref externalHyperlinks,
                ref images, ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges,
                ref unsupportedMeasurements, ref approximatedFontFamilyLists, ref unsupportedFontFamilies,
                ref mappedFields, ref unsupportedFields, handledUnsupportedFieldElements, notes, allowNotes: false);
        }
    }

    private static void CopyOdtHeaderFooterFallback(OdtHeaderFooter source, WordHeaderFooter target,
        WordOpenDocumentConversionOptions options, CultureInfo textCaseCulture,
        HashSet<XElement> handledUnsupportedFieldElements) {
        // The fallback renders the same source story in another page slot; it is not another source item.
        int hyperlinks = 0, externalHyperlinks = 0, images = 0, bookmarks = 0;
        int approximatedRuns = 0, approximatedBookmarkRanges = 0, unsupportedMeasurements = 0;
        int approximatedFontFamilyLists = 0, unsupportedFontFamilies = 0, mappedFields = 0, unsupportedFields = 0;
        var duplicateNotes = new NoteMappingStats();
        CopyOdtHeaderFooter(source, target, options, textCaseCulture, ref hyperlinks, ref externalHyperlinks,
            ref images, ref bookmarks, ref approximatedRuns, ref approximatedBookmarkRanges, ref unsupportedMeasurements,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies, ref mappedFields, ref unsupportedFields,
            handledUnsupportedFieldElements, duplicateNotes);
    }
}
