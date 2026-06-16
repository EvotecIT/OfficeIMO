using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfXmlNamespace> ReadXmlNamespaces(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? namespaceTable = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "xmlnstbl");
            if (namespaceTable == null) return Array.Empty<RtfXmlNamespace>();

            var namespaces = new List<RtfXmlNamespace>();
            foreach (RtfGroup namespaceGroup in namespaceTable.Children.OfType<RtfGroup>()) {
                RtfControlWord? control = namespaceGroup.Children.OfType<RtfControlWord>().FirstOrDefault(word => word.Name == "xmlns");
                string uri = CollectPlainText(namespaceGroup, ansiCodePage, unicodeSkipCount).Trim().TrimEnd(';').Trim();
                if (control?.Parameter is int id && id >= 0 && !string.IsNullOrWhiteSpace(uri)) {
                    namespaces.Add(new RtfXmlNamespace(id, uri));
                }
            }

            return namespaces;
        }
    }
}
