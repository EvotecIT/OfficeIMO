using System.Text;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint.OpenXmlSdk {
    /// <summary>
    /// The SDK side of <see cref="OpenXml.PowerPointXmlNode"/>: the same element operations on the SDK's elements, which
    /// it edits in place. Elements are passed as <see cref="object"/>, so callers outside the bridge name no SDK type.
    /// New elements are made through their parent, so they get the SDK's typed classes and the caller's typed access to
    /// them keeps working.
    /// </summary>
    internal static class PowerPointSdkXml {
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
            return OfficeXmlNamespaces.TryGetKnownId(uri, out int id) && id == namespaceId;
        }

        internal static object? Parent(object element) => E(element).Parent;

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

        internal static string? GetAttribute(object element, int namespaceId, string localName) {
            OpenXmlElement value = E(element);
            if (!value.HasAttributes) return null;
            string uri = Uri(namespaceId);
            foreach (OpenXmlAttribute attribute in value.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(attribute.NamespaceUri, uri, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }
            return null;
        }

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

        internal static object Clone(object element) => E(element).CloneNode(true);

        internal static byte[] OuterXml(object element) => Encoding.UTF8.GetBytes(E(element).OuterXml);

        /// <summary>Inserts an SDK element, moving it out of a detached copy, and returns it.</summary>
        internal static object InsertBefore(object parent, object child, object? reference) {
            OpenXmlElement target = E(parent), value = E(child);
            if (value.Parent != null) value = value.CloneNode(true);
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
            if (reference == null) target.AppendChild(child);
            else target.InsertBefore(child, E(reference));
            return child;
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
