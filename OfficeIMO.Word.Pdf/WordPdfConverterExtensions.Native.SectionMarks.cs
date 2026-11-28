
namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Recognizes a section terminator with no body content, while retaining its editable properties and anchors.</summary>
    internal static bool IsSectionMarkOnly(OfficeXmlNode paragraph) =>
        paragraph.WChild("pPr").WChild("sectPr").Exists &&
        paragraph.Elements().All(element => element.IsW("pPr") || element.IsW("bookmarkStart") ||
            element.IsW("bookmarkEnd") || element.IsW("proofErr"));

    /// <summary>Retains the paragraph line of an otherwise empty authored section.</summary>
    private static bool TryRenderNativeEmptySectionMark(INativePdfFlow flow, WordSection section,
        IReadOnlyList<WordElement> elements, NativeDocumentDefaults defaults, NativeFontMap fonts,
        WordPdfExportOptions? options = null) {
        OfficeXmlNode source = GetNativeEmptySectionMarkSource(section, elements);
        if (source.IsNone) return false;

        // Wrap the original paragraph without adding a run or detaching its
        // section context. The normal empty-paragraph owner resolves the mark's
        // inherited font, spacing and visibility without changing the source.
        WordParagraph mark = section._document.ParagraphView(source);
        foreach (OfficeXmlNode bookmark in source.WElements("bookmarkStart")) {
            string? name = bookmark.WAttr("name");
            if (!string.IsNullOrEmpty(name)) flow.Bookmark(name!);
        }
        RecordNativeParagraphBarTabDiagnostics(mark, options, "empty section paragraph");
        RenderNativeEmptyParagraph(flow, mark, CreateNativeParagraphStyle(mark, defaults, fonts), defaults, fonts);
        return true;
    }

    /// <summary>Shares the exact otherwise-empty section admission between paint and navigation projection.</summary>
    private static OfficeXmlNode GetNativeEmptySectionMarkSource(WordSection section, IReadOnlyList<WordElement> elements) {
        if (elements.Any(element => element is not WordParagraph paragraph ||
                !IsSectionMarkOnly(paragraph.Node()))) return default;
        OfficeXmlNode properties = section.Node().Parent;
        if (!properties.IsW("pPr")) return default;
        OfficeXmlNode source = properties.Parent;
        if (!source.IsW("p") || !IsSectionMarkOnly(source)) return default;
        return source;
    }
}
