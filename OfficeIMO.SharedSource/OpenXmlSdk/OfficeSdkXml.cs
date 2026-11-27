using System.Text;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// The SDK side of <see cref="OfficeXmlNode"/>: the same element operations on the SDK's elements, which
    /// it edits in place. Elements are passed as <see cref="object"/>, so callers outside the bridge name no SDK type.
    /// New elements are made through their parent, so they get the SDK's typed classes and the caller's typed access to
    /// them keeps working.
    /// </summary>
    internal static class OfficeSdkXml {
        private static OpenXmlElement E(object element) => (OpenXmlElement)element;

        internal static string LocalName(object element) => E(element).LocalName;

        internal static string NamespaceUri(object element) => E(element).NamespaceUri;

        internal static int NamespaceId(object element) =>
            OfficeXmlNamespaces.TryGetKnownId(E(element).NamespaceUri, out int id) ? id : OfficeXmlNamespaces.None;

        internal static bool Is(object element, int namespaceId, string localName) {
            OpenXmlElement value = E(element);
            if (!string.Equals(value.LocalName, localName, StringComparison.Ordinal)) return false;
            string uri = value.NamespaceUri;
            if (namespaceId == OfficeXmlNamespaces.None) return uri.Length == 0;
            // The transitional URI the id stands for, then any other spelling of it (strict) through the lookup.
            if (string.Equals(uri, OfficeXmlNamespaces.KnownUri(namespaceId), StringComparison.Ordinal)) return true;
            return OfficeXmlNamespaces.TryGetKnownId(uri, out int id) && id == namespaceId;
        }

        internal static object? Parent(object element) => E(element).Parent;

        /// <summary>Adds the descendant elements with this name, in document order, as the element handle walks them.</summary>
        internal static void CollectDescendants(object element, int namespaceId, string localName, List<OfficeXmlNode> found) {
            foreach (OpenXmlElement descendant in E(element).Descendants()) {
                if (descendant is OpenXmlMiscNode) continue;
                if (string.Equals(descendant.LocalName, localName, StringComparison.Ordinal) && Is(descendant, namespaceId, localName)) {
                    found.Add(OfficeXmlNode.FromSdk(descendant));
                }
            }
        }

        /// <summary>
        /// Adds the descendant elements with this name, in document order, leaving out <c>extLst</c> and
        /// <c>mc:AlternateContent</c> elements and what they hold.
        /// </summary>
        internal static void CollectContentDescendants(object element, int namespaceId, string localName, List<OfficeXmlNode> found) {
            OpenXmlElement root = E(element);
            OpenXmlElement? current = root.FirstChild;
            while (current != null) {
                bool skip = current is OpenXmlMiscNode || current.LocalName == "extLst" ||
                    current.LocalName == "AlternateContent" && Is(current, OfficeXmlNamespaces.MarkupCompatibility, "AlternateContent");
                if (!skip && string.Equals(current.LocalName, localName, StringComparison.Ordinal) && Is(current, namespaceId, localName))
                    found.Add(OfficeXmlNode.FromSdk(current));
                OpenXmlElement? next = skip ? null : current.FirstChild;
                if (next == null) {
                    for (OpenXmlElement? up = current; up != null && !ReferenceEquals(up, root); up = up.Parent) {
                        next = up.NextSibling();
                        if (next != null) break;
                    }
                }
                current = next;
            }
        }

        internal static void CollectChildNodes(object element, List<OfficeXmlNode> children) {
            foreach (OpenXmlElement child in E(element).ChildElements) children.Add(OfficeXmlNode.FromSdk(child));
        }

        internal static object? NextNode(object element) => E(element).NextSibling();

        internal static object? FirstChild(object element) => SkipForward(E(element).FirstChild);

        internal static object? LastChild(object element) => SkipBackward(E(element).LastChild);

        internal static object? NextSibling(object element) => SkipForward(E(element).NextSibling());

        internal static object? PreviousSibling(object element) => SkipBackward(E(element).PreviousSibling());

        private static OpenXmlElement? SkipForward(OpenXmlElement? element) {
            while (element is OpenXmlMiscNode) element = element.NextSibling();
            return element;
        }

        private static OpenXmlElement? SkipBackward(OpenXmlElement? element) {
            while (element is OpenXmlMiscNode) element = element.PreviousSibling();
            return element;
        }

        internal static bool IsAttached(object element) {
            OpenXmlElement value = E(element);
            while (value.Parent != null) value = value.Parent;
            return value is OpenXmlPartRootElement;
        }

        /// <summary>The SDK part whose root holds the element, or null for a detached element.</summary>
        internal static object? PartOf(object element) {
            OpenXmlElement value = E(element);
            while (value.Parent != null) value = value.Parent;
            return (value as OpenXmlPartRootElement)?.OpenXmlPart;
        }

        // Element classes and attribute names seen without the attribute. The SDK's own lookup by name doesn't copy the
        // element's attributes but throws when the attribute is missing, so it is tried only for names that have always
        // been there; the rest go through the list.
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<(Type Element, string LocalName, string Uri), bool> AttributesSeenMissing = new();

        internal static string? GetAttribute(object element, int namespaceId, string localName) {
            OpenXmlElement value = E(element);
            if (!value.HasAttributes) return null;
            string uri = Uri(namespaceId);
            var key = (value.GetType(), localName, uri);
            if (!AttributesSeenMissing.ContainsKey(key)) {
                try {
                    return value.GetAttribute(localName, uri).Value;
                } catch (KeyNotFoundException) {
                    AttributesSeenMissing.TryAdd(key, true);
                    return null;
                }
            }
            // Indexed, so reading an attribute doesn't also box an enumerator over the list the SDK builds.
            IList<OpenXmlAttribute> attributes = value.GetAttributes();
            for (int index = 0; index < attributes.Count; index++) {
                OpenXmlAttribute attribute = attributes[index];
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(attribute.NamespaceUri, uri, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }
            return null;
        }

        internal static void CollectAttributes(object element, List<OfficeXmlAttribute> attributes) {
            OpenXmlElement value = E(element);
            if (!value.HasAttributes) return;
            foreach (OpenXmlAttribute attribute in value.GetAttributes()) {
                attributes.Add(new OfficeXmlAttribute(attribute.NamespaceUri, attribute.LocalName, attribute.Value ?? string.Empty));
            }
        }

        internal static bool HasAttributes(object element) => E(element).HasAttributes;

        internal static void DeclareNamespace(object element, string prefix, string namespaceUri) => E(element).AddNamespaceDeclaration(prefix, namespaceUri);

        internal static void SetAttribute(object element, int namespaceId, string localName, string? value) {
            OpenXmlElement target = E(element);
            string uri = Uri(namespaceId);
            if (value == null) {
                if (target.HasAttributes) target.RemoveAttribute(localName, uri);
                return;
            }
            string prefix = namespaceId == OfficeXmlNamespaces.None ? string.Empty : OfficeXmlPrefixes.Canonical(namespaceId);
            target.SetAttribute(new OpenXmlAttribute(prefix, localName, uri, value));
        }

        internal static string InnerText(object element) => E(element).InnerText;

        internal static void SetInnerText(object element, string text) {
            switch (E(element)) {
                case OpenXmlLeafTextElement leaf:
                    leaf.Text = text;
                    break;
                case OpenXmlCompositeElement composite:
                    composite.RemoveAllChildren();
                    if (text.Length > 0) composite.InnerXml = Escape(text);
                    break;
                default:
                    throw new InvalidOperationException("The element '" + E(element).LocalName + "' can't hold text.");
            }
        }

        internal static void Remove(object element) => E(element).Remove();

        internal static void RemoveAllChildren(object element) => E(element).RemoveAllChildren();

        internal static object Clone(object element) => E(element).CloneNode(true);

        internal static byte[] OuterXml(object element) => Encoding.UTF8.GetBytes(E(element).OuterXml);

        // One parsed element per parent type and name. New elements are clones of it, so they get the typed class the
        // parent expects there without parsing markup each time.
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<PrototypeKey, OpenXmlElement> Prototypes = new();

        // Every new element and every insert looks a prototype up, so the key hashes the parent type and the name cheaply
        // rather than the whole namespace URI.
        private readonly struct PrototypeKey : IEquatable<PrototypeKey> {
            private readonly Type _parent;
            private readonly string _uri;
            private readonly string _localName;

            internal PrototypeKey(Type parent, string uri, string localName) {
                _parent = parent;
                _uri = uri;
                _localName = localName;
            }

            public bool Equals(PrototypeKey other) =>
                ReferenceEquals(_parent, other._parent)
                && string.Equals(_localName, other._localName, StringComparison.Ordinal)
                && (ReferenceEquals(_uri, other._uri) || string.Equals(_uri, other._uri, StringComparison.Ordinal));

            public override bool Equals(object? obj) => obj is PrototypeKey other && Equals(other);

            public override int GetHashCode() {
                int length = _localName.Length;
                int hash = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_parent);
                hash = (hash * 31) + length;
                if (length > 0) hash = (hash * 31) + _localName[0] + (_localName[length - 1] << 8) + (_localName[length >> 1] << 16);
                return (hash * 31) + _uri.Length;
            }
        }

        // Whether an element of a typed class needs reading again under a parent of a typed class, as InsertBefore decides.
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<(Type Parent, Type Child), bool> Retyped = new();

        private static OpenXmlElement Prototype(OpenXmlElement parent, string namespaceUri, string localName, string prefix) {
            var key = new PrototypeKey(parent.GetType(), namespaceUri, localName);
            if (Prototypes.TryGetValue(key, out OpenXmlElement? prototype)) return prototype;
            OpenXmlElement probe = parent.CloneNode(false);
            probe.InnerXml = prefix.Length == 0
                ? "<" + localName + " xmlns=\"" + namespaceUri + "\"/>"
                : "<" + prefix + ":" + localName + " xmlns:" + prefix + "=\"" + namespaceUri + "\"/>";
            prototype = probe.FirstChild ?? throw new InvalidOperationException("The markup holds no element.");
            prototype.Remove();
            // The parent's part declares the namespace; the element keeps none of its own, as a typed member makes it.
            if (prototype.NamespaceDeclarations.Any()) prototype.RemoveNamespaceDeclaration(prefix);
            return Prototypes.GetOrAdd(key, prototype);
        }

        /// <summary>
        /// A new, detached element with the typed class <paramref name="context"/> expects as its child, or null when the
        /// SDK has no typed class for it there.
        /// </summary>
        internal static object? CreateElement(object context, int namespaceId, string localName) {
            OpenXmlElement prototype = Prototype(E(context), Uri(namespaceId), localName, OfficeXmlPrefixes.Canonical(namespaceId));
            // A prototype has no attributes or children, so a shallow copy is the whole element.
            return prototype is OpenXmlUnknownElement ? null : prototype.CloneNode(false);
        }

        /// <summary>
        /// A new, detached element in a namespace outside the fixed ones: the typed class <paramref name="context"/> expects
        /// as its child, or an unknown element. Neither is parsed from markup, which would leave the element's namespace
        /// declaration repeated on the part root once the subtree is attached.
        /// </summary>
        internal static object CreateElement(object context, string namespaceUri, string localName, string prefix) {
            OpenXmlElement prototype = Prototype(E(context), namespaceUri, localName, prefix);
            return prototype is OpenXmlUnknownElement ? new OpenXmlUnknownElement(prefix, localName, namespaceUri) : prototype.CloneNode(false);
        }

        /// <summary>
        /// Builds the SDK element for an element of OfficeIMO's own tree from the typed classes <paramref name="parent"/>
        /// and its descendants expect, or returns null when part of it has no typed class or holds text, so the caller
        /// parses its markup instead. Typed classes keep every attribute under the name it has in the tree.
        /// </summary>
        internal static object? CreateFromTree(object parent, OfficeXmlTree tree, int node) => Build(E(parent), tree, node);

        private static OpenXmlElement? Build(OpenXmlElement parent, OfficeXmlTree tree, int node) {
            int namespaceId = tree.NamespaceOf(node);
            string? uri = OfficeXmlNamespaces.KnownUri(namespaceId);
            if (uri == null) return null;
            OpenXmlElement prototype = Prototype(parent, uri, tree.LocalNameOf(node), OfficeXmlPrefixes.Canonical(namespaceId));
            if (prototype is OpenXmlUnknownElement) return null;
            OpenXmlElement element = prototype.CloneNode(true);
            int count = tree.AttributeCountOf(node);
            for (int position = 0; position < count; position++) {
                string value = tree.AttributeAt(node, position, out int attributeNamespace, out string localName);
                if (attributeNamespace == OfficeXmlNamespaces.None) {
                    element.SetAttribute(new OpenXmlAttribute(string.Empty, localName, string.Empty, value));
                    continue;
                }
                string? attributeUri = attributeNamespace == OfficeXmlNamespaces.Xmlns ? null : OfficeXmlNamespaces.KnownUri(attributeNamespace);
                if (attributeUri == null) return null;
                element.SetAttribute(new OpenXmlAttribute(OfficeXmlPrefixes.Canonical(attributeNamespace), localName, attributeUri, value));
            }
            for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) return null;
                OpenXmlElement? built = Build(element, tree, child);
                if (built == null) return null;
                element.AppendChild(built);
            }
            return element;
        }

        /// <summary>Inserts an SDK element, moving it out of a detached copy, and returns it.</summary>
        internal static object InsertBefore(object parent, object child, object? reference) {
            OpenXmlElement target = E(parent), value = E(child);
            if (value.Parent != null) value = value.CloneNode(true);
            // An element made for another parent may have a typed class this one doesn't take; read it again here. Where
            // the SDK has no typed class for it under this parent out of context, such as under a detached mc:Choice,
            // the element stays as it is.
            if (value is not OpenXmlUnknownElement && value.NamespaceUri.Length > 0) {
                // A typed class always has the same name, so the parent and child classes decide.
                var pair = (target.GetType(), value.GetType());
                if (!Retyped.TryGetValue(pair, out bool retype)) {
                    OpenXmlElement expected = Prototype(target, value.NamespaceUri, value.LocalName, value.Prefix);
                    retype = expected is not OpenXmlUnknownElement && value.GetType() != expected.GetType();
                    Retyped.TryAdd(pair, retype);
                }
                if (retype) return InsertMarkupBefore(parent, OuterXml(value), reference);
            }
            if (reference == null) target.AppendChild(value);
            else target.InsertBefore(value, E(reference));
            return value;
        }

        /// <summary>
        /// Parses an element's markup as a child of <paramref name="parent"/>, so the SDK picks the typed class the
        /// parent expects there, and inserts it.
        /// </summary>
        internal static object InsertMarkupBefore(object parent, byte[] markup, object? reference) {
            OpenXmlElement target = E(parent);
            OpenXmlElement probe = target.CloneNode(false);
            probe.InnerXml = Encoding.UTF8.GetString(markup);
            OpenXmlElement child = probe.FirstChild ?? throw new InvalidOperationException("The markup holds no element.");
            child.Remove();
            DropInheritedDeclarations(target, child);
            OfficeXmlTree source = OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());
            RestoreAttributeNames(child, source, source.Root);
            if (reference == null) target.AppendChild(child);
            else target.InsertBefore(child, E(reference));
            return child;
        }

        // The SDK reads a few attributes under another name when it parses markup, such as w:left and w:right on w:ind as
        // w:start and w:end. Its typed members write the name the markup used, so that name is put back.
        private static void RestoreAttributeNames(OpenXmlElement element, OfficeXmlTree source, int node) {
            int count = source.AttributeCountOf(node);
            for (int position = 0; position < count && element.HasAttributes; position++) {
                string value = source.AttributeAt(node, position, out int namespaceId, out string localName);
                if (namespaceId == OfficeXmlNamespaces.Xmlns) continue;
                string? uri = namespaceId == OfficeXmlNamespaces.None ? string.Empty : OfficeXmlNamespaces.KnownUri(namespaceId);
                if (uri == null || GetAttribute(element, uri, localName) != null) continue;
                foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                    if (!string.Equals(attribute.NamespaceUri, uri, StringComparison.Ordinal)
                        || !string.Equals(attribute.Value, value, StringComparison.Ordinal)
                        || source.HasAttribute(node, namespaceId, attribute.LocalName)) continue;
                    element.RemoveAttribute(attribute.LocalName, uri);
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, localName, uri, value));
                    break;
                }
            }
            OpenXmlElement? child = SkipForward(element.FirstChild);
            for (int sourceChild = source.FirstChildOf(node); sourceChild != OfficeXmlTree.NoNode && child != null;
                sourceChild = source.NextSiblingOf(sourceChild)) {
                if (source.KindOf(sourceChild) != OfficeXmlTreeNodeKind.Element) continue;
                if (string.Equals(child.LocalName, source.LocalNameOf(sourceChild), StringComparison.Ordinal)) {
                    RestoreAttributeNames(child, source, sourceChild);
                }
                child = SkipForward(child.NextSibling());
            }
        }

        private static string? GetAttribute(OpenXmlElement element, string uri, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(attribute.NamespaceUri, uri, StringComparison.Ordinal)) return attribute.Value;
            }
            return null;
        }

        // The markup declared every prefix it uses; the ones the parent already binds the same way are dropped again.
        private static void DropInheritedDeclarations(OpenXmlElement parent, OpenXmlElement child) {
            List<KeyValuePair<string, string>>? redundant = null;
            foreach (KeyValuePair<string, string> declaration in child.NamespaceDeclarations) {
                if (string.Equals(parent.LookupNamespace(declaration.Key), declaration.Value, StringComparison.Ordinal)) {
                    (redundant ??= new List<KeyValuePair<string, string>>()).Add(declaration);
                }
            }
            if (redundant == null) return;
            foreach (KeyValuePair<string, string> declaration in redundant) child.RemoveNamespaceDeclaration(declaration.Key);
        }

        private static string Uri(int namespaceId) =>
            OfficeXmlNamespaces.KnownUri(namespaceId) ?? throw new ArgumentOutOfRangeException(nameof(namespaceId), "Only fixed namespace ids name a namespace here.");

        private static string Escape(string text) {
            var builder = new StringBuilder(text.Length + 8);
            foreach (char character in text) {
                switch (character) {
                    case '<': builder.Append("&lt;"); break;
                    case '>': builder.Append("&gt;"); break;
                    case '&': builder.Append("&amp;"); break;
                    default: builder.Append(character); break;
                }
            }
            return builder.ToString();
        }
    }
}
