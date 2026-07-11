namespace OfficeIMO.OpenDocument;

/// <summary>Enumerates logical ODF table rows without losing header-row containers.</summary>
internal static class OdfTableRowElements {
    internal static IEnumerable<XElement> Enumerate(XElement table) {
        foreach (XElement child in table.Elements()) {
            if (child.Name == OdfNamespaces.Table + "table-row") {
                yield return child;
            } else if (child.Name == OdfNamespaces.Table + "table-header-rows") {
                foreach (XElement row in child.Elements(OdfNamespaces.Table + "table-row")) yield return row;
            }
        }
    }
}
