using System.Collections.Concurrent;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// The WordprocessingML OfficeIMO writes as it is, kept in XML files embedded in the assembly: the building block
/// galleries (<c>BuildingBlocks/</c>, one <c>w:sdt</c> per template, keyed by the name of its enum value) and the
/// default parts and styles of a new document (<c>Parts/Defaults/</c>). The files are read through the owned XML layer;
/// <c>OpenXmlSdk/WordEmbeddedMarkup.OpenXmlSdk.cs</c> hands their markup to the SDK.
/// </summary>
internal static partial class WordEmbeddedMarkup {
    internal const string CoverPages = "CoverPages";
    internal const string PageNumbers = "PageNumbers";
    internal const string TablesOfContents = "TablesOfContents";
    internal const string Watermarks = "Watermarks";

    private const string BuildingBlocksPrefix = "OfficeIMO.Word.BuildingBlocks.";
    private const string DefaultsPrefix = "OfficeIMO.Word.Parts.Defaults.";

    private static readonly ConcurrentDictionary<string, Lazy<Dictionary<string, byte[]>>> Galleries = new(StringComparer.Ordinal);
    private static readonly Lazy<Dictionary<string, byte[]>> DefaultStyleMarkup = new(LoadDefaultStyles);

    /// <summary>
    /// The markup of the template <paramref name="name"/> from <paramref name="gallery"/>: one <c>w:sdt</c> that declares
    /// every prefix it uses.
    /// </summary>
    internal static byte[] GetBuildingBlockMarkup(string gallery, string name, string parameterName) {
        Dictionary<string, byte[]> templates = Galleries.GetOrAdd(gallery, key => new Lazy<Dictionary<string, byte[]>>(() => LoadGallery(key))).Value;
        return templates.TryGetValue(name, out byte[]? markup) ? markup : throw new ArgumentOutOfRangeException(parameterName);
    }

    /// <summary>
    /// The markup of an element of <c>Parts/Defaults/WordDefaultStyles.xml</c>: a <c>w:style</c> by its style id, or
    /// <c>docDefaults</c> or <c>latentStyles</c>.
    /// </summary>
    internal static byte[] GetDefaultStyleMarkup(string key) =>
        DefaultStyleMarkup.Value.TryGetValue(key, out byte[]? markup)
            ? markup
            : throw new ArgumentOutOfRangeException(nameof(key), key, "No default style element has this key.");

    /// <summary>The bytes of the default part <paramref name="fileName"/>, as a new part gets them.</summary>
    internal static Stream OpenDefaultPart(string fileName) => OpenResource(DefaultsPrefix + fileName);

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
