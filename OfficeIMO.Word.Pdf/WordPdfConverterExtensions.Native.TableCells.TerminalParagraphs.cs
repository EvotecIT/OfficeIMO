using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// A single empty final paragraph immediately after a nested table carries
    /// Word's cell-end mark rather than an additional physical paragraph line.
    /// Multiple blank paragraphs and paragraphs with their own content or layout
    /// semantics retain ordinary flow. Semantic runs and source XML stay intact.
    /// </summary>
    private static bool IsNativeNestedTableCellEndCarrier(IReadOnlyList<WordElement> elements,
        int firstIndex, int lastIndex, WordParagraph paragraph,
        Func<WordParagraph, (int Level, string Marker)?>? getMarker, NativeParagraphPaginationDefaults pagination) {
        if (pagination.KeepTogether == true || pagination.KeepWithNext == true ||
            firstIndex != lastIndex || lastIndex != elements.Count - 1 || firstIndex == 0 ||
            elements[firstIndex - 1] is not WordTable)
            return false;
        W.Paragraph source = paragraph._paragraph;
        if (source.ChildElements.Any(child => child is not W.ParagraphProperties) ||
            source.ParagraphProperties?.SectionProperties != null || HasNativePageBreakBefore(paragraph) ||
            WordParagraph.IsSectionMarkOnly(source) || getMarker?.Invoke(paragraph) != null ||
            WordDocumentTraversal.GetListInfo(paragraph) != null || GetHeadingLevel(paragraph) > 0 ||
            HasNativeParagraphBarTabs(paragraph) || HasNativeParagraphFrame(paragraph) ||
            HasNativeParagraphBorder(GetNativeEffectiveParagraphBorders(paragraph)) ||
            ParseNativeColor(GetNativeEffectiveParagraphShadingFill(paragraph)).HasValue)
            return false;
        return true;
    }
}
