using System.Xml.Linq;

namespace OfficeIMO.Word;

public partial class WordSmartArt {
    /// <summary>Gets attached editable nodes in their current diagram order.</summary>
    public IReadOnlyList<WordSmartArtNode> Nodes {
        get {
            var (document, namespaces, _) = LoadNodeParagraphs();
            return GetOrderedEditableNodePoints(document, namespaces.dgm)
                .Where(point => NodeParagraph(point, namespaces.dgm, namespaces.a) != null)
                .Select(point => new WordSmartArtNode(this, (string?)point.Attribute("modelId")
                    ?? throw new InvalidOperationException("A SmartArt node has no model identity.")))
                .ToArray();
        }
    }

    internal string ReadNodeText(string id) {
        var (document, namespaces, _) = LoadNodeParagraphs();
        XElement paragraph = FindNodeParagraph(document, namespaces.dgm, namespaces.a, id);
        return string.Concat(paragraph.Elements(namespaces.a + "r").Elements(namespaces.a + "t")
            .Select(text => (string)text));
    }

    internal void WriteNodeText(string id, string text) {
        var (document, namespaces, _, part) = LoadNodeParagraphsWithPart();
        XElement paragraph = FindNodeParagraph(document, namespaces.dgm, namespaces.a, id);
        ReplaceParagraphText(paragraph, namespaces.a, text);
        SaveDiagramData(part, document);
    }

    private static XElement FindNodeParagraph(XDocument document, XNamespace diagram, XNamespace drawing, string id) {
        XElement? point = GetEditableNodePoints(document, diagram)
            .SingleOrDefault(node => (string?)node.Attribute("modelId") == id);
        return point == null ? throw new InvalidOperationException("The SmartArt node is no longer attached to its diagram.")
            : NodeParagraph(point, diagram, drawing)
                ?? throw new InvalidOperationException("The SmartArt node has no text paragraph.");
    }

    private static XElement? NodeParagraph(XElement point, XNamespace diagram, XNamespace drawing) =>
        (point.Element(diagram + "t") ?? point.Element(diagram + "txBody"))?.Element(drawing + "p");
}
