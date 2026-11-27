using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// Reads of <see cref="OfficeXmlNode"/> elements that give the results the SDK's members give: attributes without
/// namespace declarations, text the way <c>InnerText</c> reads it, namespace lookup, and attributes in namespaces outside
/// the fixed set. Both element backings answer the same way.
/// </summary>
internal static class WordXmlNodes {
    /// <summary>
    /// The attributes of an element, without namespace declarations and markup-compatibility attributes, which the SDK
    /// keeps apart from them.
    /// </summary>
    internal static List<WordXmlAttribute> Attributes(OfficeXmlNode element) {
        var attributes = new List<WordXmlAttribute>();
        AddAttributes(element, attributes);
        return attributes;
    }

    internal static void AddAttributes(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
        if (element.IsNone) return;
        if (element.Tree is OfficeXmlTree tree) {
            int node = element.Node;
            int count = tree.AttributeCountOf(node);
            for (int index = 0; index < count; index++) {
                string value = tree.GetAttributeAt(node, index, out OfficeXmlName name);
                if (IsDeclarationOrCompatibility(name)) continue;
                attributes.Add(new WordXmlAttribute(name.Prefix ?? string.Empty, name.LocalName, NamespaceUriOf(tree, name), value));
            }
            return;
        }
        WordPackageView.SdkAttributes(element.SdkElement!, attributes);
    }

    /// <summary>Whether the element has an attribute that <paramref name="match"/> accepts, read without a list.</summary>
    internal static bool AnyAttribute(OfficeXmlNode element, Func<string, string, string, bool> match) {
        if (element.IsNone) return false;
        if (element.Tree is OfficeXmlTree tree) {
            int node = element.Node;
            int count = tree.AttributeCountOf(node);
            for (int index = 0; index < count; index++) {
                string value = tree.GetAttributeAt(node, index, out OfficeXmlName name);
                if (IsDeclarationOrCompatibility(name)) continue;
                if (match(name.LocalName, NamespaceUriOf(tree, name), value)) return true;
            }
            return false;
        }
        foreach (WordXmlAttribute attribute in Attributes(element)) {
            if (match(attribute.LocalName, attribute.NamespaceUri, attribute.Value)) return true;
        }
        return false;
    }

    private static bool IsDeclarationOrCompatibility(in OfficeXmlName name) =>
        name.NamespaceId == OfficeXmlNamespaces.Xmlns
        || name.NamespaceId == OfficeXmlNamespaces.None && string.Equals(name.LocalName, "xmlns", StringComparison.Ordinal)
        || name.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility;

    private static string NamespaceUriOf(OfficeXmlTree tree, in OfficeXmlName name) =>
        name.NamespaceId == OfficeXmlNamespaces.None ? string.Empty : tree.Namespaces.GetUriById(name.UriId);

    /// <summary>The value of an attribute named by namespace URI, for namespaces outside the fixed set.</summary>
    internal static string? Attribute(OfficeXmlNode element, string namespaceUri, string localName) {
        if (element.IsNone) return null;
        if (OfficeXmlNamespaces.TryGetKnownId(namespaceUri, out int id)) return element.Attribute(id, localName);
        if (namespaceUri.Length == 0) return element.Attribute(localName);
        foreach (WordXmlAttribute attribute in Attributes(element)) {
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                && string.Equals(attribute.NamespaceUri, namespaceUri, StringComparison.Ordinal)) {
                return attribute.Value;
            }
        }
        return null;
    }

    /// <summary>Removes an attribute named by namespace URI.</summary>
    internal static void RemoveAttribute(OfficeXmlNode element, string namespaceUri, string localName) {
        if (element.IsNone) return;
        if (element.Tree is OfficeXmlTree tree) {
            tree.RemoveAttribute(element.Node, NamespaceIdOf(tree, namespaceUri), localName);
            return;
        }
        WordPackageView.SdkRemoveAttribute(element.SdkElement!, localName, namespaceUri);
    }

    /// <summary>Sets an attribute named by namespace URI, keeping its place when it exists.</summary>
    internal static void SetAttribute(OfficeXmlNode element, string prefix, string namespaceUri, string localName, string value) {
        if (element.Tree is OfficeXmlTree tree) {
            tree.SetAttribute(element.Node, NamespaceIdOf(tree, namespaceUri), localName, value);
            return;
        }
        WordPackageView.SdkSetAttribute(element.SdkElement!, prefix, localName, namespaceUri, value);
    }

    private static int NamespaceIdOf(OfficeXmlTree tree, string namespaceUri) {
        if (namespaceUri.Length == 0) return OfficeXmlNamespaces.None;
        return OfficeXmlNamespaces.TryGetKnownId(namespaceUri, out int id) ? id : tree.Namespaces.GetId(namespaceUri);
    }

    /// <summary>The namespace a prefix is bound to at the element, or null.</summary>
    internal static string? LookupNamespace(OfficeXmlNode element, string prefix) {
        if (element.IsNone) return null;
        if (element.Tree is OfficeXmlTree tree) {
            for (int node = element.Node; node != OfficeXmlTree.NoNode; node = tree.ParentOf(node)) {
                int count = tree.AttributeCountOf(node);
                for (int index = 0; index < count; index++) {
                    string value = tree.GetAttributeAt(node, index, out OfficeXmlName name);
                    if (name.NamespaceId == OfficeXmlNamespaces.Xmlns && string.Equals(name.LocalName, prefix, StringComparison.Ordinal)) return value;
                    if (prefix.Length == 0 && name.NamespaceId == OfficeXmlNamespaces.None && string.Equals(name.LocalName, "xmlns", StringComparison.Ordinal)) return value;
                }
            }
            return null;
        }
        return WordPackageView.SdkLookupNamespace(element.SdkElement!, prefix);
    }

    /// <summary>
    /// The text of the element as the SDK's <c>InnerText</c> reads it: the text of its leaf elements, in document order.
    /// Whitespace between child elements is not text.
    /// </summary>
    internal static string InnerText(OfficeXmlNode element) {
        if (element.IsNone) return string.Empty;
        if (element.IsSdk) return element.InnerText;
        OfficeXmlNode first = element.FirstChild;
        if (first.IsNone) return element.InnerText;
        if (first.NextSibling.IsNone) {
            OfficeXmlNode grandchild = first.FirstChild;
            if (grandchild.IsNone) return first.InnerText;
        }
        var builder = new StringBuilder();
        AppendInnerText(element, builder);
        return builder.ToString();
    }

    internal static void AppendInnerText(OfficeXmlNode element, StringBuilder builder) {
        OfficeXmlNode child = element.FirstChild;
        if (child.IsNone) {
            builder.Append(element.InnerText);
            return;
        }
        for (; !child.IsNone; child = child.NextSibling) AppendInnerText(child, builder);
    }

    /// <summary>The elements below <paramref name="element"/> in document order.</summary>
    internal static IEnumerable<OfficeXmlNode> Descendants(OfficeXmlNode element) {
        if (element.IsNone) yield break;
        OfficeXmlNode node = element.FirstChild;
        while (!node.IsNone) {
            yield return node;
            OfficeXmlNode child = node.FirstChild;
            if (!child.IsNone) {
                node = child;
                continue;
            }
            while (!node.IsNone && node != element && node.NextSibling.IsNone) node = node.Parent;
            if (node.IsNone || node == element) yield break;
            node = node.NextSibling;
        }
    }

    /// <summary>The nearest ancestor with this name, the element itself excluded, or none.</summary>
    internal static OfficeXmlNode Ancestor(OfficeXmlNode element, int namespaceId, string localName) =>
        element.IsNone ? default : element.Parent.AncestorOrSelf(namespaceId, localName);

    /// <summary>The first element with this name below <paramref name="element"/>, or none.</summary>
    internal static OfficeXmlNode FirstDescendant(OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in Descendants(element)) {
            if (node.Is(namespaceId, localName)) return node;
        }
        return default;
    }

    internal static bool HasDescendant(OfficeXmlNode element, int namespaceId, string localName) => !FirstDescendant(element, namespaceId, localName).IsNone;

    /// <summary>The root of the element's tree: the part root for an attached element.</summary>
    internal static OfficeXmlNode Root(OfficeXmlNode element) {
        OfficeXmlNode node = element;
        while (!node.Parent.IsNone) node = node.Parent;
        return node;
    }
}
