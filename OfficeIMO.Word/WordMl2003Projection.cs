using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

/// <summary>Projects the supported Word 2003 XML vocabulary into ordinary editable Word parts.</summary>
internal sealed partial class WordMl2003Projection {
    private static readonly XNamespace Word = WordXmlInterchange.WordMl;
    private static readonly XNamespace Auxiliary = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private static readonly XNamespace ModernWord = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly CancellationToken _token;
    private readonly Dictionary<string, OfficeDiagnostic> _omissions = new(StringComparer.Ordinal);
    private WordMl2003Projection(OfficeOpcEditablePackage package, OfficePackageSecurityOptions policy, CancellationToken token) {
        _package = package;
        _resourcePolicy = policy;
        _token = token;
    }

    internal static IReadOnlyList<OfficeDiagnostic> WritePackage(XElement root, Stream destination,
        OfficePackageSecurityOptions policy, CancellationToken token) {
        XElement? body = root.Element(Word + "body");
        if (body == null || body.ElementsAfterSelf(Word + "body").Any())
            throw new InvalidDataException("WordML 2003 must contain exactly one Word body.");

        using var package = OfficeOpcEditablePackage.Create();
        var projection = new WordMl2003Projection(package, policy, token);
        projection.ValidateDocumentSingletons(root);
        projection.InitializeReconstructionLimits(root, policy);
        projection.AdmitSource(root, policy);
        projection.ReserveRelationship(null, createsPart: true);
        package.AddRelationship(null, OfficeOpcPackage.OfficeDocumentRelationshipType, MainPart);
        OfficeXmlTree document = package.CreateXmlPart(MainPart, WordPartContentTypePrefix + "document.main+xml", W, "document");
        projection._story = new Story(MainPart, document);
        projection.ReadDocumentParts(root);
        int output = AppendWord(document, document.Root, "body");
        int pendingSection = NoNode;
        foreach (int block in projection.ProjectBlocks(body)) {
            if (pendingSection != NoNode) {
                int paragraph = document.LastChildOf(output);
                int properties = paragraph == NoNode ? NoNode : document.FindChild(paragraph, W, "pPr");
                if (paragraph == NoNode || !document.IsElement(paragraph, W, "p") ||
                    properties != NoNode && document.FindChild(properties, W, "sectPr") != NoNode) {
                    paragraph = AppendWord(document, output, "p");
                    properties = NoNode;
                }
                if (properties == NoNode) properties = document.GetOrInsertChild(paragraph, W, "pPr", WPView.Order);
                document.InsertOrdered(properties, pendingSection, WPPrView.Order);
                pendingSection = NoNode;
            }
            if (document.IsElement(block, W, "sectPr")) pendingSection = block;
            else document.Append(output, block);
        }
        if (pendingSection != NoNode) document.Append(output, pendingSection);
        if (document.FirstChildOf(output) == NoNode) AppendWord(document, output, "p");
        token.ThrowIfCancellationRequested();
        package.Save(destination);
        return projection._omissions.Values.ToArray();
    }

    private void AdmitSource(XElement root, OfficePackageSecurityOptions policy) {
        CheckActive(policy.Macros, "macrosPresent", Word + "docSuppData", "macros");
        CheckActive(policy.EmbeddedPayloads, "embeddedObjPresent", Word + "docOleData", "embedded OLE content");
        CheckActive(policy.ActiveX, "ocxPresent", Word + "ocx", "ActiveX content");
        if (policy.ExternalRelationships == OfficePackageContentPolicy.Reject &&
            root.Descendants(Word + "hlink").Any(element => !string.IsNullOrEmpty((string?)element.Attribute(Word + "dest"))))
            throw new InvalidDataException("WordML 2003 contains a link rejected by the external relationship policy.");

        void CheckActive(OfficePackageContentPolicy contentPolicy, string flag, XName elementName, string label) {
            XAttribute? attribute = root.Attribute(Word + flag);
            if ((attribute == null || !IsTrue(attribute.Value)) && !root.Descendants(elementName).Any()) return;
            if (contentPolicy == OfficePackageContentPolicy.Reject)
                throw new InvalidDataException("WordML 2003 contains " + label + " rejected by the package content policy.");
            if (attribute != null && IsTrue(attribute.Value)) Omit(attribute.Name, label);
        }
    }

    // Word 2003 XML projects the legacy vocabulary only. Extensions and alternate branches
    // are opaque to the projection, including when they contain familiar resource declarations.
    private static IEnumerable<XElement> SourceDescendants(XElement root) {
        XElement? node = root.Elements().FirstOrDefault();
        while (node != null) {
            XElement? child = null;
            if (node.Name.LocalName != "extLst" && node.Name != MarkupCompatibility + "AlternateContent") {
                yield return node;
                child = node.Elements().FirstOrDefault();
            }
            if (child != null) {
                node = child;
                continue;
            }
            while (node != root && node != null && NextElement(node) == null) node = node.Parent;
            node = node == root || node == null ? null : NextElement(node);
        }

        static XElement? NextElement(XElement element) {
            XNode? sibling = element.NextNode;
            while (sibling != null && sibling is not XElement) sibling = sibling.NextNode;
            return (XElement?)sibling;
        }
    }

    private static readonly XNamespace MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private IEnumerable<int> ProjectBlocks(XElement container, bool inTableCell = false) {
        CheckAttributes(container, container.Name == Word + "hdr" || container.Name == Word + "ftr" ? "type" :
            container.Name == Word + "footnote" || container.Name == Word + "endnote" ? "type suppressRef" : "");
        foreach (XElement child in container.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "p") yield return ProjectParagraph(child, inTableCell);
            else if (child.Name == Word + "tbl") yield return ProjectTable(child);
            else if (child.Name == Auxiliary + "sect" || child.Name == Auxiliary + "sub-section") {
                foreach (int block in ProjectBlocks(child, inTableCell)) yield return block;
            } else if (child.Name == Word + "sectPr") {
                if (inTableCell || _story.PartName != MainPart)
                    Omit(child, inTableCell ? "section properties inside a table cell" : "section properties inside a header or footer");
                else yield return ProjectSectionProperties(child);
            } else if (!inTableCell || child.Name != Word + "tcPr") {
                Omit(child, "block content");
            }
        }
    }

    private int ProjectParagraph(XElement source, bool inTableCell) {
        CheckAttributes(source, "");
        OfficeXmlTree tree = _story.Tree;
        int paragraph = tree.CreateElement(W, "p");
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "pPr") {
                int properties = ProjectProperties(tree, child, "pPr", parent: "p");
                if (child.Element(Word + "sectPr") is { } section) {
                    if (inTableCell || _story.PartName != MainPart)
                        Omit(section, inTableCell ? "section properties inside a table cell" : "section properties inside a header or footer");
                    else tree.InsertOrdered(properties, ProjectSectionProperties(section), WPPrView.Order);
                }
                if (tree.FirstChildOf(properties) != NoNode) ReplaceWordChild(tree, paragraph, properties, WPView.Order);
            } else AppendInline(paragraph, child);
        }
        return paragraph;
    }

    private void AppendInline(int target, XElement source) {
        _token.ThrowIfCancellationRequested();
        if (source.Name == Word + "hlink") {
            AppendHyperlink(target, source);
        } else if (source.Name == Word + "r") {
            CheckAttributes(source, "");
            OfficeXmlTree tree = _story.Tree;
            int run = tree.CreateElement(W, "r");
            foreach (XElement child in source.Elements()) {
                if (child.Name == Word + "t" && !child.HasElements) {
                    CheckAttributes(child, "");
                    AppendText(tree, run, "t", child.Value);
                } else if (child.Name == Word + "tab") {
                    CheckAttributes(child, "");
                    AppendWord(tree, run, "tab");
                } else if (child.Name == Word + "br" || child.Name == Word + "fldChar") {
                    XElement? mapped = MapProperties(child, child.Name.LocalName);
                    if (mapped != null) AppendMapped(tree, run, mapped);
                } else if (child.Name == Word + "rPr") {
                    ReplaceWordChild(tree, run, ProjectProperties(tree, child, "rPr"), WRView.Order);
                } else if (child.Name == Word + "footnote" || child.Name == Word + "endnote") {
                    AppendNote(run, child);
                } else if (child.Name == Word + "footnoteRef") {
                    AppendNoteMark(run, child);
                } else if (child.Name == Word + "pict") {
                    XElement? picture = ProjectPicture(child);
                    if (picture != null) AppendMapped(tree, run, picture);
                } else if (child.Name == Word + "instrText" && !child.HasElements) {
                    CheckAttributes(child, "");
                    AppendText(tree, run, "instrText", child.Value);
                } else Omit(child, "run content");
            }
            tree.Append(target, run);
        } else Omit(source, "inline content");
    }

    private int ProjectTable(XElement source) {
        CheckAttributes(source, "");
        OfficeXmlTree tree = _story.Tree;
        XElement? sourceGrid = source.Element(Word + "tblGrid");
        int columns = AdmitTableGrid(source, sourceGrid);
        int table = tree.CreateElement(W, "tbl");
        XElement? properties = source.Element(Word + "tblPr");
        tree.Append(table, properties == null ? tree.CreateElement(W, "tblPr") : ProjectProperties(tree, properties, "tblPr"));
        int grid = sourceGrid == null ? tree.CreateElement(W, "tblGrid") : ProjectProperties(tree, sourceGrid, "tblGrid");
        tree.Append(table, grid);
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "tblPr" || child.Name == Word + "tblGrid") continue;
            if (child.Name != Word + "tr") { Omit(child, "table content"); continue; }
            CheckAttributes(child, "");
            int row = tree.CreateElement(W, "tr");
            bool hasCell = false;
            foreach (XElement cell in child.Elements()) {
                if (cell.Name == Word + "trPr" || cell.Name == Word + "tblPrEx") {
                    tree.Append(row, ProjectProperties(tree, cell, cell.Name.LocalName));
                } else if (cell.Name == Word + "tc") {
                    int outputCell = AppendWord(tree, row, "tc");
                    if (cell.Element(Word + "tcPr") is { } cellProperties) {
                        int mapped = ProjectProperties(tree, cellProperties, "tcPr");
                        tree.Append(outputCell, mapped);
                        int span = tree.FindChild(mapped, W, "gridSpan");
                        if (span != NoNode && OfficeXmlViewValues.ReadInt32(tree, span, W, "val") == 0)
                            tree.SetAttribute(span, W, "val", "1");
                    }
                    foreach (int block in ProjectBlocks(cell, inTableCell: true)) tree.Append(outputCell, block);
                    int last = tree.LastChildOf(outputCell);
                    if (last == NoNode || !tree.IsElement(last, W, "p")) AppendWord(tree, outputCell, "p");
                    hasCell = true;
                } else Omit(cell, "row content");
            }
            if (hasCell) tree.Append(table, row);
            else Omit(child, "table row without supported cells");
        }
        if (tree.FindChild(table, W, "tr") == NoNode) {
            Omit(source, "table without supported cells");
            int row = AppendWord(tree, table, "tr");
            int cell = AppendWord(tree, row, "tc");
            AppendWord(tree, cell, "p");
        }
        if (tree.FirstChildOf(grid) == NoNode) {
            for (int i = 0; i < columns; i++) AppendWord(tree, grid, "gridCol");
        }
        return table;
    }

    private int ProjectSectionProperties(XElement source) {
        var stories = new HashSet<(XName Name, string? Type)>();
        foreach (XElement child in source.Elements().Where(element => element.Name == Word + "hdr" || element.Name == Word + "ftr")) {
            if (!stories.Add((child.Name, (string?)child.Attribute(Word + "type"))))
                throw new InvalidDataException("WordML 2003 contains a duplicate header or footer type in a section.");
        }
        int properties = ProjectProperties(_story.Tree, source, "sectPr");
        foreach (XElement child in source.Elements().Where(element => element.Name == Word + "hdr" || element.Name == Word + "ftr"))
            AppendHeaderFooter(properties, child);
        return properties;
    }

    private void Omit(XElement element, string reason) => Omit(element.Name, reason);

    private void Omit(XName name, string reason) {
        string identity = name + "|" + reason;
        if (_omissions.ContainsKey(identity)) return;
        _omissions.Add(identity, new OfficeDiagnostic("WORDML2003_OMITTED", OfficeDiagnosticSeverity.Warning,
            "WordML 2003 " + reason + " was omitted (" + name.LocalName + ").",
            location: name.ToString(), format: "WordML 2003", lossKind: OfficeConversionLossKind.Omission));
    }

    private static bool IsTrue(string value) => value is "1" or "on" or "true" or "yes";
}
