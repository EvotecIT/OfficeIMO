using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

internal static partial class WordFieldInventory {
    private static readonly char[] CompatibilityListSeparators = { ' ', '\t', '\r', '\n' };

    // Walk the story in document order without entering opaque markup.
    internal static IEnumerable<OfficeXmlNode> EnumerateFieldContent(OfficeXmlNode root, bool reserveInactiveCompatibilityIds = false) {
        const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
        if (!reserveInactiveCompatibilityIds && IsIgnorableMarkup(root, out bool processRoot)) {
            if (!processRoot) yield break;
        } else if (root.LocalName == "extLst") {
            yield break;
        }
        if (!reserveInactiveCompatibilityIds && root.Is(Mc, "AlternateContent"))
            root = WordAlternateContentResolver.SelectBranch(root);
        OfficeXmlNode node = root.FirstChild;
        while (node.Exists) {
            OfficeXmlNode next = default;
            if (!reserveInactiveCompatibilityIds && node.Is(Mc, "AlternateContent")) {
                next = WordAlternateContentResolver.SelectBranch(node).FirstChild;
            } else if (!reserveInactiveCompatibilityIds && IsIgnorableMarkup(node, out bool processContent)) {
                if (processContent) next = node.FirstChild;
            } else if (node.LocalName != "extLst") {
                yield return node;
                next = node.FirstChild;
            }

            if (next.IsNone) {
                for (OfficeXmlNode up = node; up.Exists && up != root; up = up.Parent) {
                    // Returning from a selected branch resumes after AlternateContent.
                    if (!reserveInactiveCompatibilityIds && (up.Is(Mc, "Choice") || up.Is(Mc, "Fallback"))) continue;
                    next = up.NextSibling;
                    if (next.Exists) break;
                }
            }
            node = next;
        }
    }

    private static bool IsIgnorableMarkup(OfficeXmlNode node, out bool processContent) {
        processContent = false;
        if (node.NamespaceId != OfficeXmlNamespaces.None) return false;
        string namespaceUri = node.NamespaceUri;
        if (namespaceUri.Length == 0 || WordAlternateContentResolver.IsSupportedExtensionNamespace(namespaceUri)) return false;

        bool ignorable = false;
        for (OfficeXmlNode scope = node; scope.Exists; scope = scope.Parent) {
            if (!ignorable) ignorable = MatchesCompatibilityDirective(scope, "Ignorable", namespaceUri, null);
            if (!processContent) processContent = MatchesCompatibilityDirective(scope, "ProcessContent", namespaceUri, node.LocalName);
            if (ignorable && processContent) break;
        }
        return ignorable;
    }

    private static bool MatchesCompatibilityDirective(OfficeXmlNode scope, string directive, string namespaceUri, string? localName) {
        string? names = scope.Attribute(OfficeXmlNamespaces.MarkupCompatibility, directive);
        if (string.IsNullOrEmpty(names)) return false;
        foreach (string name in names!.Split(CompatibilityListSeparators, StringSplitOptions.RemoveEmptyEntries)) {
            string prefix = name;
            if (localName != null) {
                int separator = name.IndexOf(':');
                if (separator <= 0) continue;
                string target = name.Substring(separator + 1);
                if (target != "*" && target != localName) continue;
                prefix = name.Substring(0, separator);
            }

            // Resolve the directive where it was declared, even if the prefix was rebound below it.
            if (WordXmlNodes.LookupNamespace(scope, prefix) == namespaceUri) return true;
        }
        return false;
    }
}
