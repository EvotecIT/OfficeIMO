using System.Collections.Concurrent;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// The WordprocessingML OfficeIMO writes as it is, kept in XML files embedded in the assembly: the building block
/// galleries (<c>BuildingBlocks/</c>, one <c>w:sdt</c> per template, keyed by the name of its enum value) and the
/// default parts and styles of a new document (<c>Parts/Defaults/</c>). The files are read through the owned XML layer;
/// the SDK gets elements built from their markup, or a part's bytes.
/// </summary>
internal static class WordEmbeddedMarkup {
    internal const string CoverPages = "CoverPages";
    internal const string PageNumbers = "PageNumbers";
    internal const string TablesOfContents = "TablesOfContents";
    internal const string Watermarks = "Watermarks";

    private const string BuildingBlocksPrefix = "OfficeIMO.Word.BuildingBlocks.";
    private const string DefaultsPrefix = "OfficeIMO.Word.Parts.Defaults.";

    private static readonly ConcurrentDictionary<string, Lazy<Dictionary<string, byte[]>>> Galleries = new(StringComparer.Ordinal);
    private static readonly ConcurrentDictionary<(string Gallery, string Name), Lazy<SdtBlock>> BuildingBlocks = new();
    private static readonly Lazy<Dictionary<string, byte[]>> DefaultStyleMarkup = new(LoadDefaultStyles);
    private static readonly ConcurrentDictionary<string, Lazy<OpenXmlElement>> DefaultStyleElements = new(StringComparer.Ordinal);

    /// <summary>
    /// A new, detached copy of the template <paramref name="name"/> from <paramref name="gallery"/>, ready for the caller
    /// to give it identifiers and insert it.
    /// </summary>
    internal static SdtBlock CreateBuildingBlock(string gallery, string name, string parameterName) {
        Dictionary<string, byte[]> templates = Galleries.GetOrAdd(gallery, key => new Lazy<Dictionary<string, byte[]>>(() => LoadGallery(key))).Value;
        if (!templates.TryGetValue(name, out byte[]? markup)) throw new ArgumentOutOfRangeException(parameterName);
        // Each template is parsed once, even when threads ask for it together; callers get copies they can edit.
        SdtBlock template = BuildingBlocks.GetOrAdd((gallery, name), _ => new Lazy<SdtBlock>(() => Parse(markup, xml => new SdtBlock(xml)))).Value;
        return (SdtBlock)template.CloneNode(true);
    }

    /// <summary>
    /// A new copy of an element of <c>Parts/Defaults/WordDefaultStyles.xml</c>: a <c>w:style</c> by its style id, or
    /// <c>docDefaults</c> or <c>latentStyles</c>.
    /// </summary>
    internal static T CreateDefaultStyleElement<T>(string key, Func<string, T> create) where T : OpenXmlElement {
        // Each element is parsed once, even when threads ask for it together; callers get copies they can edit.
        OpenXmlElement element = DefaultStyleElements.GetOrAdd(key, name => new Lazy<OpenXmlElement>(() => {
            if (!DefaultStyleMarkup.Value.TryGetValue(name, out byte[]? markup)) throw new ArgumentOutOfRangeException(nameof(key), name, "No default style element has this key.");
            return Parse(markup, create);
        })).Value;
        return (T)element.CloneNode(true);
    }

    /// <summary>Writes the default part <paramref name="fileName"/> into <paramref name="part"/>.</summary>
    internal static void FeedDefaultPart(OpenXmlPart part, string fileName) {
        using Stream stream = OpenResource(DefaultsPrefix + fileName);
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
            tree.GetAttributeAt(node, index, out string prefix, out string localName, out string namespaceUri, out string value, out bool declaration);
            if (declaration) continue;
            (written ??= new List<OpenXmlAttribute>(count)).Add(new OpenXmlAttribute(prefix, localName, namespaceUri, value));
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

    private static Dictionary<string, byte[]> LoadGallery(string gallery) {
        OfficeXmlTree tree = ReadTree(BuildingBlocksPrefix + gallery + ".xml");
        var templates = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (!tree.IsElement(node, OfficeXmlNamespaces.None, "template")) continue;
            int block = tree.FindChild(node, OfficeXmlNamespaces.WordprocessingML, "sdt");
            templates[tree.GetAttribute(node, OfficeXmlNamespaces.None, "name")!] = tree.WriteElement(block);
        }
        return templates;
    }

    private static Dictionary<string, byte[]> LoadDefaultStyles() {
        OfficeXmlTree tree = ReadTree(DefaultsPrefix + "WordDefaultStyles.xml");
        const int W = OfficeXmlNamespaces.WordprocessingML;
        var elements = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        for (int node = tree.FirstChildOf(tree.Root); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
            if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) continue;
            string key = tree.IsElement(node, W, "style") ? tree.GetAttribute(node, W, "styleId")! : tree.LocalNameOf(node);
            elements[key] = tree.WriteElement(node);
        }
        return elements;
    }

    private static OfficeXmlTree ReadTree(string resource) {
        byte[] bytes;
        using (Stream stream = OpenResource(resource)) {
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            bytes = buffer.ToArray();
        }
        return OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
    }

    private static Stream OpenResource(string resource) =>
        typeof(WordEmbeddedMarkup).Assembly.GetManifestResourceStream(resource)
        ?? throw new InvalidOperationException($"The embedded markup {resource} is missing from the assembly.");
}
