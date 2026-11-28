using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordTableOfContent {
    internal static WordTableOfContent? FindFirstTable(WordDocument document) {
        if (document._ownedModel is { BodyHasFieldsOrContentControls: false }) return null;
        OfficeXmlNode body = WordFieldInventory.BodyNode(document);
        foreach (OfficeXmlNode control in body.Children(TocNamespace, "sdt")) {
            if (IsTocGallery(control)) return WrapImportedTable(document, control);
        }
        return WrapBareTable(document, body);
    }

    private static WordTableOfContent? WrapImportedTable(WordDocument document, OfficeXmlNode control) =>
        control.IsSdk ? WrapSdkControl(document, control) : new WordTableOfContent(document, control);

    private static WordTableOfContent? WrapBareTable(WordDocument document, OfficeXmlNode body) {
        var paragraphs = body.Children(TocNamespace, "p").ToList();
        foreach (OfficeXmlNode paragraph in paragraphs) {
            int? simpleField = FindSimpleTocOrIndexFieldChildIndex(paragraph);
            if (simpleField.HasValue) return WrapSimpleField(document, body, paragraph, simpleField.Value);
        }

        for (int index = 0; index < paragraphs.Count; index++) {
            int? start = FindTocOrIndexComplexFieldStart(paragraphs[index]);
            if (start.HasValue && FindComplexFieldEnd(paragraphs, index, start.Value) is { } end)
                return WrapComplexField(document, body, paragraphs, index, start.Value, end);
        }
        return null;
    }

    private static WordTableOfContent? WrapSimpleField(WordDocument document, OfficeXmlNode body, OfficeXmlNode paragraph, int index) {
        int position = GetChildPosition(body, paragraph);
        OfficeXmlNode control = CreateImportedControl(document, body);
        OfficeXmlNode prefix = SplitFieldStartPrefix(paragraph, index);
        if (prefix.Exists) index = FindSimpleTocOrIndexFieldChildIndex(paragraph) ?? index;

        OfficeXmlNode content = control.Child(TocNamespace, "sdtContent");
        OfficeXmlNode field = SplitFieldTail(paragraph, index, complex: false);
        if (field.IsNone) {
            paragraph.Remove();
            field = paragraph;
        }
        content.Append(field);
        return InsertImportedTable(document, body, control, prefix, position);
    }

    private static WordTableOfContent? WrapComplexField(WordDocument document, OfficeXmlNode body, List<OfficeXmlNode> paragraphs,
        int index, int start, (int Paragraph, int Child) end) {
        OfficeXmlNode paragraph = paragraphs[index];
        int position = GetChildPosition(body, paragraph);
        OfficeXmlNode control = CreateImportedControl(document, body);
        OfficeXmlNode prefix = SplitFieldStartPrefix(paragraph, start);
        if (prefix.Exists) end = FindComplexFieldEnd(paragraphs, index, 0) ?? end;

        OfficeXmlNode content = control.Child(TocNamespace, "sdtContent");
        for (int current = index; current <= end.Paragraph; current++) {
            OfficeXmlNode source = paragraphs[current];
            OfficeXmlNode field = current == end.Paragraph ? SplitFieldTail(source, end.Child, complex: true) : default;
            if (field.IsNone) {
                source.Remove();
                field = source;
            }
            content.Append(field);
        }
        return InsertImportedTable(document, body, control, prefix, position);
    }

    private static int GetChildPosition(OfficeXmlNode body, OfficeXmlNode child) {
        int position = 0;
        for (OfficeXmlNode sibling = body.FirstChild; sibling.Exists && sibling != child; sibling = sibling.NextSibling)
            position++;
        return position;
    }

    private static WordTableOfContent? InsertImportedTable(WordDocument document, OfficeXmlNode body, OfficeXmlNode control,
        OfficeXmlNode prefix, int position) {
        document.AssignNewSdtIds(control.Child(TocNamespace, "sdtContent"));
        OfficeXmlNode reference = body.FirstChild;
        for (int index = 0; index < position && reference.Exists; index++) reference = reference.NextSibling;
        if (prefix.Exists) body.InsertBefore(prefix, reference);
        body.InsertBefore(control, reference);
        document._ownedModel?.MarkChanged();
        return WrapImportedTable(document, control);
    }

    private static OfficeXmlNode CreateImportedControl(WordDocument document, OfficeXmlNode body) {
        OfficeXmlNode control = body.CreateElement(TocNamespace, "sdt");
        OfficeXmlNode properties = AppendTocChild(control, "sdtPr");
        AppendTocChild(properties, "id");
        OfficeXmlNode gallery = AppendTocChild(properties, "docPartObj");
        AppendTocChild(gallery, "docPartGallery").SetAttribute(TocNamespace, "val", "Table of Contents");
        AppendTocChild(gallery, "docPartUnique");
        AppendTocChild(control, "sdtContent");
        // Reserve the wrapper identifier before any paragraph is split or detached.
        document.AssignNewSdtIds(control);
        return control;
    }

    private static int? FindTocOrIndexComplexFieldStart(OfficeXmlNode paragraph) {
        const int W = TocNamespace;
        var fieldStarts = new Stack<(int Index, System.Text.StringBuilder Instruction)>();
        int index = 0;
        foreach (OfficeXmlNode child in paragraph.Children()) {
            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(child)) {
                if (IsRawFieldMarker(element, "begin")) {
                    fieldStarts.Push((index, new System.Text.StringBuilder()));
                } else if (IsRawFieldMarker(element, "end") && fieldStarts.Count > 0) {
                    fieldStarts.Pop();
                } else if (element.Is(W, "instrText") && fieldStarts.Count > 0) {
                    var current = fieldStarts.Peek();
                    current.Instruction.Append(element.InnerText);
                    if (IsRawTocOrIndexInstruction(current.Instruction.ToString())) return current.Index;
                }
            }
            index++;
        }
        return null;
    }

    private static int? FindSimpleTocOrIndexFieldChildIndex(OfficeXmlNode paragraph) {
        const int W = TocNamespace;
        int index = 0;
        foreach (OfficeXmlNode child in paragraph.Children()) {
            if (child.Is(W, "fldSimple") && IsRawTocOrIndexInstruction(child.Attribute(W, "instr"))) return index;
            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(child)) {
                if (element.Is(W, "fldSimple") && IsRawTocOrIndexInstruction(element.Attribute(W, "instr"))) return index;
            }
            index++;
        }
        return null;
    }

    private static (int Paragraph, int Child)? FindComplexFieldEnd(IReadOnlyList<OfficeXmlNode> paragraphs, int start, int firstChild) {
        int depth = 0;
        for (int row = start; row < paragraphs.Count; row++) {
            int index = -1;
            foreach (OfficeXmlNode child in paragraphs[row].Children()) {
                index++;
                if (row == start && index < firstChild) continue;
                foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(child)) {
                    if (IsRawFieldMarker(element, "begin")) depth++;
                    else if (depth > 0 && IsRawFieldMarker(element, "end") && --depth == 0) return (row, index);
                }
            }
        }
        return null;
    }

    private static bool IsRawTocOrIndexInstruction(string? instruction) =>
        !string.IsNullOrWhiteSpace(instruction) && GetFieldType(instruction!) is WordFieldType.TOC or WordFieldType.Index;

    private static bool IsRawFieldMarker(OfficeXmlNode node, string type) =>
        node.Is(TocNamespace, "fldChar") && node.Attribute(TocNamespace, "fldCharType") == type;

    private static OfficeXmlNode SplitFieldStartPrefix(OfficeXmlNode paragraph, int index) {
        const int W = TocNamespace;
        var children = paragraph.Children().ToList();
        OfficeXmlNode field = children[index];
        OfficeXmlNode sameRun = default;
        if (field.Is(W, "r")) {
            OfficeXmlNode candidate = paragraph.CreateElement(W, "r");
            foreach (OfficeXmlNode child in field.Children()) {
                if (IsRawFieldMarker(child, "begin")) break;
                candidate.Append(child.Clone());
            }
            if (HasRawFieldContent(candidate)) sameRun = candidate;
        }
        if (sameRun.IsNone && !children.Take(index).Any(HasRawFieldContent)) return default;

        OfficeXmlNode prefix = CreateSplitParagraph(paragraph);
        for (int i = 0; i < index; i++) {
            if (children[i].Is(W, "pPr")) continue;
            prefix.Append(children[i].Clone());
            children[i].Remove();
        }
        if (sameRun.Exists) prefix.Append(sameRun);
        if (field.Is(W, "r")) {
            OfficeXmlNode child = field.FirstChild;
            while (child.Exists && !IsRawFieldMarker(child, "begin")) {
                OfficeXmlNode next = child.NextSibling;
                child.Remove();
                child = next;
            }
        }
        return prefix;
    }

    private static OfficeXmlNode SplitFieldTail(OfficeXmlNode paragraph, int index, bool complex) {
        const int W = TocNamespace;
        var children = paragraph.Children().ToList();
        OfficeXmlNode last = children[index];
        if (!(children.Skip(index + 1).Any(HasRawFieldContent) || complex && HasRawFieldContentAfterEnd(last)))
            return default;

        OfficeXmlNode field = CreateSplitParagraph(paragraph);
        for (int i = 0; i <= index; i++) {
            OfficeXmlNode child = children[i];
            if (child.Is(W, "pPr")) continue;
            if (complex && i == index && child.Is(W, "r")) {
                OfficeXmlNode run = AppendTocChild(field, "r");
                bool foundEnd = false;
                foreach (OfficeXmlNode runChild in child.Children().ToList()) {
                    run.Append(runChild.Clone());
                    runChild.Remove();
                    if (IsRawFieldMarker(runChild, "end")) {
                        foundEnd = true;
                        break;
                    }
                }
                if (foundEnd && HasRawFieldContent(child)) continue;
            } else {
                field.Append(child.Clone());
            }
            child.Remove();
        }
        return field;
    }

    private static OfficeXmlNode CreateSplitParagraph(OfficeXmlNode paragraph) {
        OfficeXmlNode split = paragraph.CreateElement(TocNamespace, "p");
        OfficeXmlNode properties = paragraph.Child(TocNamespace, "pPr");
        if (properties.Exists) split.Append(properties.Clone());
        return split;
    }

    private static bool HasRawFieldContent(OfficeXmlNode element) =>
        element.IsExtensionMarkup() || ContainsOpaqueRunContent(element) ||
        WordFieldInventory.EnumerateFieldContent(element).Any(IsRawFieldContent);

    private static bool ContainsOpaqueRunContent(OfficeXmlNode element) {
        foreach (OfficeXmlNode child in element.Children()) {
            if (child.IsExtensionMarkup()) return true;
        }
        return false;
    }

    private static bool HasRawFieldContentAfterEnd(OfficeXmlNode element) {
        bool afterDirectEnd = false;
        foreach (OfficeXmlNode child in element.Children()) {
            if (afterDirectEnd && child.IsExtensionMarkup()) return true;
            if (IsRawFieldMarker(child, "end")) afterDirectEnd = true;
        }

        bool afterEnd = false;
        foreach (OfficeXmlNode child in WordFieldInventory.EnumerateFieldContent(element)) {
            if (afterEnd && IsRawFieldContent(child)) return true;
            if (IsRawFieldMarker(child, "end")) afterEnd = true;
        }
        return false;
    }

    private static bool IsRawFieldContent(OfficeXmlNode node) => node.NamespaceId == TocNamespace &&
        (node.LocalName is "fldSimple" or "instrText" or "drawing" or "pict" ||
         node.LocalName == "t" && !string.IsNullOrWhiteSpace(node.InnerText));
}
