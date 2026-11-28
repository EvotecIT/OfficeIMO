using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Small helpers for reading and writing PresentationML and DrawingML through <see cref="OfficeXmlNode"/>: namespace
    /// ids, typed attribute values parsed as the schema defines them, element creation, and a markup digest that does not
    /// depend on how the part was serialized.
    /// </summary>
    internal static class LegacyPptXml {
        internal const int P = OfficeXmlNamespaces.PresentationML;
        internal const int A = OfficeXmlNamespaces.DrawingML;
        internal const int R = OfficeXmlNamespaces.Relationships;
        internal const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
        internal const int Pic = OfficeXmlNamespaces.Picture;
        internal const int C = OfficeXmlNamespaces.Chart;
        internal const int Dgm = OfficeXmlNamespaces.Diagram;

        internal const string P14Uri = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        internal const string P15Uri = "http://schemas.microsoft.com/office/powerpoint/2012/main";
        internal const string A14Uri = "http://schemas.microsoft.com/office/drawing/2010/main";

        /// <summary>An <c>xsd:boolean</c> attribute: <c>true</c>/<c>1</c> or <c>false</c>/<c>0</c>; null when absent or invalid.</summary>
        internal static bool? Bool(this OfficeXmlNode element, string localName) => ParseBoolean(element.Attribute(localName));

        internal static bool Bool(this OfficeXmlNode element, string localName, bool defaultValue) => element.Bool(localName) ?? defaultValue;

        // xsd:boolean and the integer types collapse whitespace, so surrounding XML whitespace is not part of the value.
        private static readonly char[] XmlWhitespace = { ' ', '\t', '\r', '\n' };

        private static string? Collapse(string? value) => value?.Trim(XmlWhitespace);

        internal static bool? ParseBoolean(string? value) => Collapse(value) switch {
            "true" or "1" => true,
            "false" or "0" => false,
            _ => null
        };

        internal static int? Int(this OfficeXmlNode element, string localName) => ParseInt32(element.Attribute(localName));

        internal static uint? UInt(this OfficeXmlNode element, string localName) => ParseUInt32(element.Attribute(localName));

        internal static long? Long(this OfficeXmlNode element, string localName) => ParseInt64(element.Attribute(localName));

        internal static int? ParseInt32(string? value) =>
            Collapse(value) is string text && int.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;

        internal static uint? ParseUInt32(string? value) =>
            Collapse(value) is string text && uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) ? parsed : null;

        internal static long? ParseInt64(string? value) =>
            Collapse(value) is string text && long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long parsed) ? parsed : null;

        internal static string Format(long value) => value.ToString(CultureInfo.InvariantCulture);

        internal static string Format(bool value) => value ? "1" : "0";

        /// <summary>The <c>r:id</c>-style attribute in the relationships namespace.</summary>
        internal static string? RelationshipId(this OfficeXmlNode element, string localName = "id") => element.Attribute(R, localName);

        /// <summary>The first child with this name in PresentationML.</summary>
        internal static OfficeXmlNode PChild(this OfficeXmlNode element, string localName) => element.Child(P, localName);

        /// <summary>The first child with this name in DrawingML.</summary>
        internal static OfficeXmlNode AChild(this OfficeXmlNode element, string localName) => element.Child(A, localName);

        /// <summary>Walks a path of children in one namespace; none as soon as one is missing.</summary>
        internal static OfficeXmlNode Path(this OfficeXmlNode element, int namespaceId, params string[] localNames) {
            OfficeXmlNode current = element;
            foreach (string localName in localNames) {
                if (current.IsNone) return default;
                current = current.Child(namespaceId, localName);
            }
            return current;
        }

        /// <summary>
        /// Elements below <paramref name="element"/> in document order, at any depth, extension content included, as the
        /// SDK's untyped <c>Descendants()</c> walked them.
        /// </summary>
        internal static IEnumerable<OfficeXmlNode> EveryDescendant(this OfficeXmlNode element) =>
            Walk(element, skipExtensions: false);

        /// <summary>
        /// The descendants with this name, in document order, without the content of extension lists (<c>extLst</c>).
        /// It replaces the SDK's typed <c>Descendants&lt;T&gt;()</c>, which did not see markup inside extension
        /// metadata, so a name that also appears there must not match it.
        /// </summary>
        internal static List<OfficeXmlNode> SchemaDescendants(this OfficeXmlNode element, int namespaceId, string localName) {
            var result = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode candidate in Walk(element, skipExtensions: true)) {
                if (candidate.Is(namespaceId, localName)) result.Add(candidate);
            }
            return result;
        }

        /// <summary>The first descendant with this name outside extension lists, or none; see <see cref="SchemaDescendants"/>.</summary>
        internal static OfficeXmlNode FirstSchemaDescendant(this OfficeXmlNode element, int namespaceId, string localName) {
            foreach (OfficeXmlNode candidate in Walk(element, skipExtensions: true)) {
                if (candidate.Is(namespaceId, localName)) return candidate;
            }
            return default;
        }

        private static IEnumerable<OfficeXmlNode> Walk(OfficeXmlNode element, bool skipExtensions) {
            if (element.IsNone) yield break;
            var pending = new Stack<OfficeXmlNode>();
            for (OfficeXmlNode child = element.LastChild; !child.IsNone; child = child.PreviousSibling) pending.Push(child);
            while (pending.Count > 0) {
                OfficeXmlNode current = pending.Pop();
                if (skipExtensions && IsExtensionList(current)) continue;
                yield return current;
                for (OfficeXmlNode child = current.LastChild; !child.IsNone; child = child.PreviousSibling) pending.Push(child);
            }
        }

        private static bool IsExtensionList(OfficeXmlNode element) =>
            string.Equals(element.LocalName, "extLst", StringComparison.Ordinal);

        /// <summary>Whether an ancestor (not the element itself) has this name.</summary>
        internal static bool HasAncestor(this OfficeXmlNode element, int namespaceId, string localName) =>
            !element.IsNone && !element.Parent.AncestorOrSelf(namespaceId, localName).IsNone;

        /// <summary>
        /// A SHA-256 digest of an element's markup that depends only on names, attributes (namespace declarations left
        /// out, in name order) and text, so the same content digests the same however it was written.
        /// </summary>
        internal static string Digest(OfficeXmlNode element) {
            var builder = new StringBuilder();
            AppendCanonical(element, builder);
            return DigestText(builder.ToString());
        }

        internal static string DigestText(string text) {
            byte[] bytes = Encoding.UTF8.GetBytes(text);
#if NET5_0_OR_GREATER
            byte[] hash = SHA256.HashData(bytes);
#else
            byte[] hash;
            using (SHA256 sha = SHA256.Create()) hash = sha.ComputeHash(bytes);
#endif
            var hex = new StringBuilder(hash.Length * 2);
            foreach (byte value in hash) hex.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            return hex.ToString();
        }

        /// <summary>Writes the canonical form <see cref="Digest"/> hashes.</summary>
        internal static void AppendCanonical(OfficeXmlNode element, StringBuilder builder) {
            if (element.IsNone) return;
            if (element.Tree is not { } tree) throw new InvalidOperationException("Canonical markup is read from Open XML layer trees.");
            AppendCanonical(tree, element.Node, builder);
        }

        private static void AppendCanonical(OfficeXmlTree tree, int node, StringBuilder builder) {
            if (tree.KindOf(node) == OfficeXmlTreeNodeKind.Text) {
                builder.Append('"');
                AppendEscaped(builder, tree.GetText(node));
                builder.Append('"');
                return;
            }
            builder.Append('<');
            AppendNamespace(builder, tree, tree.NamespaceOf(node));
            builder.Append(tree.LocalNameOf(node));
            int count = tree.AttributeCountOf(node);
            if (count > 0) {
                var attributes = new List<(string Name, string Value)>(count);
                for (int index = 0; index < count; index++) {
                    tree.GetAttributeAt(node, index, out int namespaceId, out string localName, out string value);
                    if (namespaceId == OfficeXmlNamespaces.Xmlns
                        || namespaceId == OfficeXmlNamespaces.None && string.Equals(localName, "xmlns", StringComparison.Ordinal)) continue;
                    var name = new StringBuilder();
                    AppendNamespace(name, tree, namespaceId);
                    name.Append(localName);
                    if (IsCompatibilityPrefixList(tree, node, namespaceId, localName)) {
                        value = ResolveCompatibilityPrefixes(tree, node, value);
                    }
                    attributes.Add((name.ToString(), value));
                }
                attributes.Sort((left, right) => string.CompareOrdinal(left.Name, right.Name));
                foreach ((string name, string value) in attributes) {
                    builder.Append(' ').Append(name).Append("=\"");
                    AppendEscaped(builder, value);
                    builder.Append('"');
                }
            }
            builder.Append('>');
            bool hasElements = false;
            for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) {
                    hasElements = true;
                    break;
                }
            }
            for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                // Indentation between elements is formatting, which serializers add or drop.
                if (hasElements && tree.KindOf(child) == OfficeXmlTreeNodeKind.Text && IsXmlWhitespace(tree.GetText(child))
                    && !PreservesSpace(tree, node)) continue;
                AppendCanonical(tree, child, builder);
            }
            builder.Append("</>");
        }

        // Markup-compatibility attributes whose values name namespaces by prefix: mc:Ignorable, mc:MustUnderstand,
        // mc:ProcessContent, mc:PreserveElements, mc:PreserveAttributes, and Requires on mc:Choice.
        private static bool IsCompatibilityPrefixList(OfficeXmlTree tree, int node, int namespaceId, string localName) {
            if (namespaceId == Mc) {
                return localName is "Ignorable" or "MustUnderstand" or "ProcessContent" or "PreserveElements" or "PreserveAttributes";
            }
            return namespaceId == OfficeXmlNamespaces.None
                && string.Equals(localName, "Requires", StringComparison.Ordinal)
                && tree.NamespaceOf(node) == Mc
                && string.Equals(tree.LocalNameOf(node), "Choice", StringComparison.Ordinal);
        }

        // Replaces each prefix in a compatibility list with the namespace it is bound to where the attribute sits, so a
        // changed binding changes the digest and a renamed prefix with the same binding doesn't.
        private static string ResolveCompatibilityPrefixes(OfficeXmlTree tree, int node, string value) {
            var resolved = new StringBuilder(value.Length);
            foreach (string token in value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                int colon = token.IndexOf(':');
                string prefix = colon < 0 ? token : token.Substring(0, colon);
                if (resolved.Length > 0) resolved.Append(' ');
                resolved.Append('{').Append(LookupNamespace(tree, node, prefix) ?? prefix).Append('}');
                if (colon >= 0) resolved.Append(token, colon + 1, token.Length - colon - 1);
            }
            return resolved.ToString();
        }

        // Only XML whitespace (space, tab, CR, LF) is indentation; other Unicode spaces such as U+00A0 are text.
        private static bool IsXmlWhitespace(string text) {
            foreach (char character in text) {
                if (character is not (' ' or '\t' or '\r' or '\n')) return false;
            }
            return true;
        }

        // Whether the nearest xml:space on the element or its ancestors is "preserve".
        private static bool PreservesSpace(OfficeXmlTree tree, int node) {
            for (int current = node; current != OfficeXmlTree.NoNode; current = tree.ParentOf(current)) {
                if (tree.KindOf(current) != OfficeXmlTreeNodeKind.Element) continue;
                int count = tree.AttributeCountOf(current);
                for (int index = 0; index < count; index++) {
                    tree.GetAttributeAt(current, index, out int namespaceId, out string localName, out string value);
                    if (namespaceId == OfficeXmlNamespaces.Xml && string.Equals(localName, "space", StringComparison.Ordinal)) {
                        return string.Equals(value, "preserve", StringComparison.Ordinal);
                    }
                }
            }
            return false;
        }

        private static string? LookupNamespace(OfficeXmlTree tree, int node, string prefix) {
            for (int current = node; current != OfficeXmlTree.NoNode; current = tree.ParentOf(current)) {
                if (tree.KindOf(current) != OfficeXmlTreeNodeKind.Element) continue;
                int count = tree.AttributeCountOf(current);
                for (int index = 0; index < count; index++) {
                    tree.GetAttributeAt(current, index, out int namespaceId, out string localName, out string uri);
                    if (namespaceId == OfficeXmlNamespaces.Xmlns && string.Equals(localName, prefix, StringComparison.Ordinal)) return uri;
                }
            }
            return null;
        }

        private static void AppendNamespace(StringBuilder builder, OfficeXmlTree tree, int namespaceId) {
            if (namespaceId == OfficeXmlNamespaces.None) return;
            builder.Append('{');
            if (namespaceId < OfficeXmlNamespaces.FirstDynamic) builder.Append(namespaceId.ToString(CultureInfo.InvariantCulture));
            else builder.Append(tree.Namespaces.GetUri(namespaceId));
            builder.Append('}');
        }

        private static void AppendEscaped(StringBuilder builder, string value) {
            foreach (char character in value) {
                if (character is '"' or '\\') builder.Append('\\');
                builder.Append(character);
            }
        }
    }
}
