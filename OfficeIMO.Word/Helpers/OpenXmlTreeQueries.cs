using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Ancestor and descendant checks that follow parent and sibling links directly. The Open XML SDK's
    /// Ancestors and Descendants allocate an iterator on every call, which adds up when every paragraph asks.
    /// </summary>
    internal static class OpenXmlTreeQueries {
        internal static T? FindAncestor<T>(OpenXmlElement element) where T : OpenXmlElement {
            for (OpenXmlElement? parent = element.Parent; parent != null; parent = parent.Parent) {
                if (parent is T match) return match;
            }
            return null;
        }

        internal static bool HasAncestor<T>(OpenXmlElement element) where T : OpenXmlElement => FindAncestor<T>(element) != null;

        internal static bool HasDescendant<T>(OpenXmlElement root) where T : OpenXmlElement {
            OpenXmlElement? node = root.FirstChild;
            while (node != null) {
                if (node is T) return true;
                OpenXmlElement? next = node.FirstChild;
                while (next == null && node != null) {
                    next = node.NextSibling();
                    if (next == null) {
                        node = node.Parent;
                        if (ReferenceEquals(node, root)) return false;
                    }
                }
                node = next;
            }
            return false;
        }
    }
}
