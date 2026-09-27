namespace OfficeIMO.OpenDocument;

/// <summary>Enumerates logical ODF table rows through row containers in document order.</summary>
internal static class OdfTableRowElements {
    internal static IEnumerable<XElement> Enumerate(XElement table) {
        foreach (XElement child in table.Elements()) {
            foreach (XElement row in EnumerateChild(child)) yield return row;
        }
    }

    internal static IEnumerable<XElement> EnumerateAfter(XElement table, XElement row) {
        for (XElement? current = row; current != null && !ReferenceEquals(current, table); current = current.Parent) {
            foreach (XElement sibling in current.ElementsAfterSelf()) {
                foreach (XElement following in EnumerateChild(sibling)) yield return following;
            }
        }
    }

    private static IEnumerable<XElement> EnumerateChild(XElement child) {
        if (child.Name == OdfNamespaces.Table + "table-row") {
            yield return child;
        } else if (child.Name == OdfNamespaces.Table + "table-header-rows"
            || child.Name == OdfNamespaces.Table + "table-rows"
            || child.Name == OdfNamespaces.Table + "table-row-group") {
            foreach (XElement nested in child.Elements()) {
                foreach (XElement row in EnumerateChild(nested)) yield return row;
            }
        }
    }
}
