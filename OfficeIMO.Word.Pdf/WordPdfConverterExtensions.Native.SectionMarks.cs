using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Retains the paragraph line of an otherwise empty authored section.</summary>
    private static bool TryRenderNativeEmptySectionMark(INativePdfFlow flow, WordSection section,
        IReadOnlyList<WordElement> elements, NativeDocumentDefaults defaults, NativeFontMap fonts,
        WordPdfExportOptions? options = null) {
        if (elements.Any(element => element is not WordParagraph paragraph ||
                !WordParagraph.IsSectionMarkOnly(paragraph._paragraph)) ||
            section._sectionProperties.Parent is not W.ParagraphProperties properties ||
            properties.Parent is not W.Paragraph source || !WordParagraph.IsSectionMarkOnly(source)) return false;

        // Wrap the original paragraph without adding a run or detaching its
        // section context. The normal empty-paragraph owner resolves the mark's
        // inherited font, spacing and visibility without changing the source.
        var mark = new WordParagraph(section._document, source);
        foreach (W.BookmarkStart bookmark in source.Elements<W.BookmarkStart>()) {
            if (!string.IsNullOrEmpty(bookmark.Name?.Value)) flow.Bookmark(bookmark.Name!.Value!);
        }
        RecordNativeParagraphBarTabDiagnostics(mark, options, "empty section paragraph");
        RenderNativeEmptyParagraph(flow, mark, CreateNativeParagraphStyle(mark, defaults, fonts), defaults, fonts);
        return true;
    }
}
