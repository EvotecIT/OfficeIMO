using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// Body paragraphs of an owned document as <see cref="WordParagraph"/> objects, split into runs the way the SDK path
/// splits them. Paragraphs holding content the owned layer doesn't model yet (hyperlinks, fields, content controls,
/// equations) make it decline, and the SDK path enumerates the document instead.
/// </summary>
internal static class WordOwnedParagraphs {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const int NoNode = OfficeXmlTree.NoNode;

    internal static bool TryCollect(WordDocument document, WordPackageModel model, out List<WordParagraph> paragraphs) {
        OfficeXmlTree tree = model.Document;
        int count = 0;
        for (int child = tree.FirstChildOf(model.Body); child != NoNode; child = tree.NextSiblingOf(child)) count++;
        paragraphs = new List<WordParagraph>(count);
        for (int child = tree.FirstChildOf(model.Body); child != NoNode; child = tree.NextSiblingOf(child)) {
            if (!tree.IsElement(child, W, "p")) continue;
            if (IsPureSectionBreak(tree, child)) continue;
            if (!TryAppendParagraph(document, tree, child, paragraphs)) return false;
        }
        return true;
    }

    private static bool IsPureSectionBreak(OfficeXmlTree tree, int paragraph) {
        int properties = tree.FindChild(paragraph, W, "pPr");
        if (properties == NoNode || tree.FindChild(properties, W, "sectPr") == NoNode) return false;
        for (int child = tree.FirstChildOf(paragraph); child != NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && child != properties) return false;
        }
        for (int child = tree.FirstChildOf(properties); child != NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && !tree.IsElement(child, W, "sectPr")) return false;
        }
        return true;
    }

    private static bool TryAppendParagraph(WordDocument document, OfficeXmlTree tree, int paragraph, List<WordParagraph> list) {
        int first = NextElement(tree, tree.FirstChildOf(paragraph));
        if (first == NoNode) {
            list.Add(WordParagraph.CreateOwned(document, tree, paragraph));
            return true;
        }
        int second = NextElement(tree, tree.NextSiblingOf(first));
        if (second == NoNode) {
            if (tree.IsElement(first, W, "r")) {
                list.Add(WordParagraph.CreateOwned(document, tree, paragraph, first));
                return true;
            }
            if (tree.IsElement(first, W, "pPr")) {
                list.Add(WordParagraph.CreateOwned(document, tree, paragraph));
                return true;
            }
        } else if (tree.IsElement(first, W, "pPr") && tree.IsElement(second, W, "r")
            && NextElement(tree, tree.NextSiblingOf(second)) == NoNode) {
            list.Add(WordParagraph.CreateOwned(document, tree, paragraph, second));
            return true;
        }

        for (int child = first; child != NoNode; child = NextElement(tree, tree.NextSiblingOf(child))) {
            if (!TryProcess(document, tree, paragraph, child, list)) return false;
        }
        return true;
    }

    private static bool TryProcess(WordDocument document, OfficeXmlTree tree, int paragraph, int element, List<WordParagraph> list) {
        int namespaceId = tree.NamespaceOf(element);
        if (namespaceId == OfficeXmlNamespaces.Math) {
            string mathName = tree.LocalNameOf(element);
            // Equations become equation paragraphs on the SDK path.
            return mathName != "oMath" && mathName != "oMathPara";
        }
        if (namespaceId != W) return true;
        switch (tree.LocalNameOf(element)) {
            case "r":
                // A field spans runs and becomes one paragraph on the SDK path.
                if (tree.FindChild(element, W, "fldChar") != NoNode) return false;
                list.Add(WordParagraph.CreateOwned(document, tree, paragraph, element));
                return true;
            case "ins":
            case "moveTo":
            case "customXml":
                for (int child = NextElement(tree, tree.FirstChildOf(element)); child != NoNode; child = NextElement(tree, tree.NextSiblingOf(child))) {
                    if (!TryProcess(document, tree, paragraph, child, list)) return false;
                }
                return true;
            case "bookmarkStart":
                list.Add(WordParagraph.CreateOwned(document, tree, paragraph, bookmark: element));
                return true;
            case "hyperlink":
            case "fldSimple":
            case "sdt":
                return false;
            default:
                // Deleted runs, range markers, proofing marks and paragraph properties add no paragraph.
                return true;
        }
    }

    private static int NextElement(OfficeXmlTree tree, int node) {
        while (node != NoNode && tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) node = tree.NextSiblingOf(node);
        return node;
    }

    /// <summary>
    /// Appends a paragraph to the body before its final section properties, as <see cref="WordDocument.AddParagraph(string)"/>
    /// does on the SDK path: <c>w:pPr</c> first, then a run with the text when there is text.
    /// </summary>
    internal static WordParagraph AddParagraph(WordDocument document, WordPackageModel model, bool withRun) {
        WordOwnedEdits.PrepareBodyEdit(model);
        OfficeXmlTree tree = model.Document;
        int body = model.Body;
        int paragraph = tree.CreateElement(W, "p");
        tree.Append(paragraph, tree.CreateElement(W, "pPr"));
        int run = NoNode;
        if (withRun) {
            run = tree.CreateElement(W, "r");
            tree.Append(paragraph, run);
        }
        int boundary = FinalSectionProperties(tree, body);
        tree.InsertBefore(body, paragraph, boundary);
        model.MarkChanged();
        return WordParagraph.CreateOwned(document, tree, paragraph, run);
    }

    // The final section properties, moved to the end of the body first when something follows them.
    private static int FinalSectionProperties(OfficeXmlTree tree, int body) {
        int last = WordOwnedSave.LastElementChild(tree, body);
        if (last != NoNode && tree.IsElement(last, W, "sectPr")) return last;
        int sectPr = NoNode;
        for (int child = tree.FirstChildOf(body); child != NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.IsElement(child, W, "sectPr")) sectPr = child;
        }
        if (sectPr == NoNode) return NoNode;
        tree.Remove(sectPr);
        tree.Append(body, sectPr);
        return sectPr;
    }
}
