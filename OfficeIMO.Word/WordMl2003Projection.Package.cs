using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const int R = OfficeXmlNamespaces.Relationships;
    private const int NoNode = OfficeXmlTree.NoNode;
    private const string MainPart = "/word/document.xml";
    private const string WordPartContentTypePrefix = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
    private const string CorePropertiesPart = "/docProps/core.xml";
    private const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";
    private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    private readonly OfficeOpcEditablePackage _package;
    private Story _story;
    private int _headerCount;
    private int _footerCount;
    private readonly Dictionary<string, int> _imageCounts = new(StringComparer.Ordinal);

    private readonly struct Story(string partName, OfficeXmlTree tree) {
        internal string PartName { get; } = partName;
        internal OfficeXmlTree Tree { get; } = tree;
    }

    private OfficeXmlTree AddMainPart(string name, string root, out string id, string? partType = null) {
        string fileName = name + ".xml";
        partType ??= name;
        ReserveRelationship(MainPart, createsPart: true);
        id = _package.AddRelationship(MainPart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + partType, fileName);
        return _package.CreateXmlPart("/word/" + fileName, WordPartContentTypePrefix + partType + "+xml", W, root);
    }

    private static int AppendWord(OfficeXmlTree tree, int parent, string name) {
        int node = tree.CreateElement(W, name);
        tree.Append(parent, node);
        return node;
    }

    private static int AppendText(OfficeXmlTree tree, int parent, string name, string text) {
        int node = AppendWord(tree, parent, name);
        tree.SetAttribute(node, OfficeXmlNamespaces.Xml, "space", "preserve");
        tree.SetInnerText(node, text);
        return node;
    }

    private static void AppendMapped(OfficeXmlTree tree, int parent, XElement source) =>
        tree.Append(parent, ImportMapped(tree, source));

    // Only projected elements reach this method. Their depth is bounded by source admission.
    private static int ImportMapped(OfficeXmlTree tree, XElement source) {
        int ns = GetNamespaceId(tree, source.Name.Namespace);
        int node = tree.CreateElement(ns, source.Name.LocalName);
        foreach (XAttribute attribute in source.Attributes()) {
            if (attribute.IsNamespaceDeclaration) {
                string prefix = attribute.Name.Namespace == XNamespace.Xmlns ? attribute.Name.LocalName : string.Empty;
                tree.DeclareNamespace(node, prefix, attribute.Value);
                continue;
            }
            int attributeNs = GetNamespaceId(tree, attribute.Name.Namespace);
            tree.SetAttribute(node, attributeNs, attribute.Name.LocalName, attribute.Value);
        }
        foreach (XElement child in source.Elements()) AppendMapped(tree, node, child);
        return node;
    }

    private static int GetNamespaceId(OfficeXmlTree tree, XNamespace ns) =>
        ns == XNamespace.None ? OfficeXmlNamespaces.None : tree.Namespaces.GetId(ns.NamespaceName);

    private static void ReplaceWordChild(OfficeXmlTree tree, int parent, int child, OfficeXmlChildOrder order) {
        int previous = tree.FindChild(parent, W, tree.LocalNameOf(child));
        if (previous != NoNode) tree.Remove(previous);
        tree.InsertOrdered(parent, child, order);
    }

    private string AddImage(byte[] bytes, string contentType) {
        string extension = contentType switch {
            "image/jpeg" => "jpg",
            "image/tiff" => "tiff",
            "image/x-icon" or "image/vnd.microsoft.icon" => "ico",
            "image/svg+xml" => "svg",
            "image/x-wmf" => "wmf",
            "image/x-emf" => "emf",
            "image/x-ms-bmp" => "bmp",
            _ => contentType.Substring(contentType.IndexOf('/') + 1)
        };
        _imageCounts.TryGetValue(extension, out int count);
        _imageCounts[extension] = ++count;
        string path = "/media/image" + (count == 1 ? "" : count.ToString(CultureInfo.InvariantCulture)) + "." + extension;
        _package.CreateBinaryPart(path, contentType, bytes);
        string target = "../media/" + path.Substring("/media/".Length);
        return _package.AddRelationship(_story.PartName, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "image", target);
    }
}
