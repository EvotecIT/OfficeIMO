using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// Body paragraphs of an owned document as <see cref="WordParagraph"/> objects, split into runs the way the SDK path
/// splits them. Paragraphs holding content the owned layer doesn't model yet (hyperlinks, fields, content controls,
/// equations) make it decline, and the SDK path enumerates the document instead.
/// </summary>
internal static class WordOwnedParagraphs {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    internal static bool IsRunningStory(OfficeXmlTree? tree) => tree != null &&
        (tree.IsElement(tree.Root, W, "hdr") || tree.IsElement(tree.Root, W, "ftr"));

    internal static bool HasFootnotes(WordPackageModel model) {
        var body = new OfficeXmlNode(model.Document, model.Body);
        if (WordFieldInventory.EnumerateFieldContent(body).Any(node => node.Is(W, "footnoteReference"))) return true;
        const string relationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "footnotes";
        foreach (string name in WordOwnedSave.RelatedParts(model, relationship)) {
            OfficeXmlTree tree = model.Package.GetXml(name);
            var notes = new OfficeXmlNode(tree, tree.Root);
            foreach (OfficeXmlNode note in notes.Children(W, "footnote")) {
                if (!note.HasAttribute(W, "type")) return true;
            }
        }
        return false;
    }
    private const int NoNode = OfficeXmlTree.NoNode;

    internal static bool TryCollect(WordDocument document, WordPackageModel model, out List<WordParagraph> paragraphs) {
        return TryCollectStory(document, model.Document, model.Body, out paragraphs);
    }

    internal static bool TryCollectStory(WordDocument document, OfficeXmlTree tree, int root, out List<WordParagraph> paragraphs) {
        return TryCollectRegion(document, tree, tree.FirstChildOf(root), NoNode, out paragraphs);
    }

    internal static bool TryCollectRegion(WordDocument document, OfficeXmlTree tree, int first, int stop,
        out List<WordParagraph> paragraphs) {
        int count = 0;
        for (int child = first; child != NoNode && child != stop; child = tree.NextSiblingOf(child)) count++;
        paragraphs = new List<WordParagraph>(count);
        for (int child = first; child != NoNode && child != stop; child = tree.NextSiblingOf(child)) {
            if (!tree.IsElement(child, W, "p")) continue;
            if (IsPureSectionBreak(tree, child)) continue;
            if (!TryAppendParagraph(document, tree, child, paragraphs)) return false;
        }
        return true;
    }

    // Page galleries and simple running stories use flat marker runs. Project their field
    // paragraphs without opening the SDK package.
    internal static List<WordParagraph> ProjectFlatFieldParagraphs(WordDocument document, OfficeXmlNode content) {
        var paragraphs = new List<WordParagraph>();
        foreach (OfficeXmlNode paragraph in content.Children(W, "p")) {
            List<OfficeXmlNode>? fieldRuns = null;
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (child.Is(W, "fldSimple")) {
                    paragraphs.Add(WordParagraph.CreateOwnedField(document, paragraph, child, runs: null));
                } else if (child.Is(W, "r")) {
                    string? marker = child.Child(W, "fldChar").Attribute(W, "fldCharType");
                    if (fieldRuns != null) {
                        fieldRuns.Add(child);
                        if (marker == "end") {
                            paragraphs.Add(WordParagraph.CreateOwnedField(document, paragraph, default, fieldRuns));
                            fieldRuns = null;
                        }
                    } else if (marker == "begin" &&
                        (paragraph.FirstChild != child || child.NextSibling.Exists)) {
                        fieldRuns = new List<OfficeXmlNode> { child };
                    } else {
                        paragraphs.Add(WordParagraph.CreateOwned(document, paragraph.Tree!, paragraph.Node, child.Node));
                    }
                }
            }
        }
        return paragraphs;
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
        int current = element;
        while (current != NoNode) {
            int nested = NoNode;
            int namespaceId = tree.NamespaceOf(current);
            string name = tree.LocalNameOf(current);
            if (namespaceId == OfficeXmlNamespaces.Math && name is "oMath" or "oMathPara") return false;
            if (namespaceId == W) {
                switch (name) {
                    case "r":
                        // A field spans runs and becomes one paragraph on the SDK path.
                        if (tree.FindChild(current, W, "fldChar") != NoNode) return false;
                        list.Add(WordParagraph.CreateOwned(document, tree, paragraph, current));
                        break;
                    case "ins":
                    case "moveTo":
                    case "customXml":
                        nested = NextElement(tree, tree.FirstChildOf(current));
                        break;
                    case "bookmarkStart":
                        list.Add(WordParagraph.CreateOwned(document, tree, paragraph, bookmark: current));
                        break;
                    case "hyperlink":
                    case "fldSimple":
                    case "sdt":
                        return false;
                }
            }

            if (nested != NoNode) {
                current = nested;
                continue;
            }
            nested = NoNode;
            for (int ancestor = current; ancestor != element; ancestor = tree.ParentOf(ancestor)) {
                nested = NextElement(tree, tree.NextSiblingOf(ancestor));
                if (nested != NoNode) break;
            }
            current = nested;
        }
        return true;
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
        return AddParagraph(document, model, withRun, FinalSectionProperties(model.Document, model.Body));
    }

    internal static WordParagraph AddParagraph(WordDocument document, WordPackageModel model, bool withRun, int boundary) {
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
