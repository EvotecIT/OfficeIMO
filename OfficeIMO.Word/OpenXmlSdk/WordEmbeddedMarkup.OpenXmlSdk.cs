using System.Collections.Concurrent;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

// Hands the embedded markup to the SDK while cover pages, page numbers, tables of contents, watermarks, styles and new
// document parts are still built as SDK elements and parts. It goes when those callers write through the owned layer.
internal static partial class WordEmbeddedMarkup {
    private static readonly ConcurrentDictionary<(string Gallery, string Name), Lazy<SdtBlock>> BuildingBlocks = new();
    private static readonly ConcurrentDictionary<string, Lazy<OpenXmlElement>> DefaultStyleElements = new(StringComparer.Ordinal);

    /// <summary>
    /// A new, detached copy of the template <paramref name="name"/> from <paramref name="gallery"/>, ready for the caller
    /// to give it identifiers and insert it.
    /// </summary>
    internal static SdtBlock CreateBuildingBlock(string gallery, string name, string parameterName) {
        byte[] markup = GetBuildingBlockMarkup(gallery, name, parameterName);
        // Each template is parsed once, even when threads ask for it together; callers get copies they can edit.
        SdtBlock template = BuildingBlocks.GetOrAdd((gallery, name), _ => new Lazy<SdtBlock>(() => Parse(markup, xml => new SdtBlock(xml)))).Value;
        return (SdtBlock)template.CloneNode(true);
    }

    /// <summary>A new copy of a default style element (see <see cref="GetDefaultStyleMarkup"/>).</summary>
    internal static T CreateDefaultStyleElement<T>(string key, Func<string, T> create) where T : OpenXmlElement {
        // Each element is parsed once, even when threads ask for it together; callers get copies they can edit.
        OpenXmlElement element = DefaultStyleElements.GetOrAdd(key, name =>
            new Lazy<OpenXmlElement>(() => Parse(GetDefaultStyleMarkup(name), create))).Value;
        return (T)element.CloneNode(true);
    }

    /// <summary>Writes the default part <paramref name="fileName"/> into <paramref name="part"/>.</summary>
    internal static void FeedDefaultPart(OpenXmlPart part, string fileName) {
        using Stream stream = OpenDefaultPart(fileName);
        part.FeedData(stream);
    }

    // Builds the SDK element from markup that declares every prefix it uses. Parsing markup on its own, the SDK reads a
    // few attributes for Office 2007: w:ind's w:left and w:right come back as w:start and w:end, and w:tblLook keeps only
    // w:val. Elements whose attributes came back different get them again exactly as written, and the root drops its
    // declarations so that, inserted, it takes them from the part as an element built in code does.
    private static T Parse<T>(byte[] markup, Func<string, T> create) where T : OpenXmlElement {
        T element = create(Encoding.UTF8.GetString(markup));
        OfficeXmlTree tree = OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());
        RestoreAttributes(tree, tree.Root, element);
        foreach (KeyValuePair<string, string> declaration in element.NamespaceDeclarations.ToList()) {
            element.RemoveNamespaceDeclaration(declaration.Key);
        }
        return element;
    }

    private static void RestoreAttributes(OfficeXmlTree tree, int node, OpenXmlElement element) {
        int count = tree.AttributeCountOf(node);
        List<OpenXmlAttribute>? written = null;
        for (int index = 0; index < count; index++) {
            OfficeXmlName name = tree.AttributeNameAt(node, index);
            if (name.NamespaceId == OfficeXmlNamespaces.Xmlns) continue;
            string namespaceUri = name.NamespaceId == OfficeXmlNamespaces.None ? string.Empty : tree.Namespaces.GetUriById(name.UriId);
            (written ??= new List<OpenXmlAttribute>(count)).Add(
                new OpenXmlAttribute(name.Prefix ?? string.Empty, name.LocalName, namespaceUri, tree.AttributeValueAt(node, index)));
        }
        if (!SameAttributes(element.GetAttributes(), written)) {
            element.ClearAllAttributes();
            if (written != null) {
                foreach (OpenXmlAttribute attribute in written) element.SetAttribute(attribute);
            }
        }

        OpenXmlElement? child = element.FirstChild;
        for (int childNode = tree.FirstChildOf(node); childNode != OfficeXmlTree.NoNode; childNode = tree.NextSiblingOf(childNode)) {
            if (tree.KindOf(childNode) != OfficeXmlTreeNodeKind.Element) continue;
            while (child is OpenXmlMiscNode) child = child.NextSibling();
            if (child == null) throw new InvalidOperationException("The embedded markup and the elements parsed from it differ.");
            RestoreAttributes(tree, childNode, child);
            child = child.NextSibling();
        }
    }

    private static bool SameAttributes(IList<OpenXmlAttribute> parsed, List<OpenXmlAttribute>? written) {
        if (parsed.Count != (written?.Count ?? 0)) return false;
        if (written == null) return true;
        foreach (OpenXmlAttribute attribute in written) {
            bool found = false;
            foreach (OpenXmlAttribute candidate in parsed) {
                if (candidate.LocalName == attribute.LocalName && candidate.NamespaceUri == attribute.NamespaceUri) {
                    found = candidate.Value == attribute.Value;
                    break;
                }
            }
            if (!found) return false;
        }
        return true;
    }
}
