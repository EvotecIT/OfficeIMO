using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

internal static partial class WordFieldUpdater {
    private static int EstimateTotalPages(WordDocument document) {
        OfficeXmlNode body = WordFieldInventory.BodyNode(document);
        int page = 1;
        foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(body)) {
            if (!node.Is(FieldNamespace, "p")) continue;
            page += CountParagraphPageBreaks(node);
            if (StartsNewPage(node)) page++;
        }
        return page;
    }

    private static int? EstimatePageForBodyField(WordDocument document, OfficeXmlNode anchor) {
        OfficeXmlNode body = WordFieldInventory.BodyNode(document);
        OfficeXmlNode target = anchor.AncestorOrSelf(FieldNamespace, "p");
        if (body.IsNone || target.IsNone) return null;

        int page = 1;
        foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(body)) {
            if (!node.Is(FieldNamespace, "p")) continue;
            if (node == target) return page + CountParagraphPageBreaks(node, anchor);
            page += CountParagraphPageBreaks(node);
            if (StartsNewPage(node)) page++;
        }
        return null;
    }

    private static int CountParagraphPageBreaks(OfficeXmlNode paragraph, OfficeXmlNode stop = default) {
        const int W = FieldNamespace;
        OfficeXmlNode pageBreakBefore = paragraph.Child(W, "pPr").Child(W, "pageBreakBefore");
        int count = pageBreakBefore.Exists && pageBreakBefore.Attribute(W, "val") is null or "true" or "1" or "on" ? 1 : 0;
        foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(paragraph)) {
            if (node == stop) break;
            if (node.Is(W, "br") && node.Attribute(W, "type") == "page") count++;
        }
        return count;
    }

    private static bool StartsNewPage(OfficeXmlNode paragraph) {
        const int W = FieldNamespace;
        OfficeXmlNode section = paragraph.Child(W, "pPr").Child(W, "sectPr");
        return section.Exists && section.Child(W, "type").Attribute(W, "val") is null or "nextPage" or "oddPage" or "evenPage";
    }
}
