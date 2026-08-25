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

    internal static IEnumerable<XElement> EnumerateAfter(XElement table, XElement row) {
        XElement? parent = row.Parent;
        if (parent == null) yield break;

        if (parent.Name == OdfNamespaces.Table + "table-header-rows") {
            foreach (XElement sibling in row.ElementsAfterSelf(OdfNamespaces.Table + "table-row")) {
                yield return sibling;
            }
            foreach (XElement child in parent.ElementsAfterSelf()) {
                foreach (XElement following in EnumerateChild(child)) yield return following;
            }
            yield break;
        }

        if (!ReferenceEquals(parent, table)) yield break;
        foreach (XElement child in row.ElementsAfterSelf()) {
            foreach (XElement following in EnumerateChild(child)) yield return following;
        }
    }

    private static IEnumerable<XElement> EnumerateChild(XElement child) {
        if (child.Name == OdfNamespaces.Table + "table-row") {
            yield return child;
        } else if (child.Name == OdfNamespaces.Table + "table-header-rows") {
            foreach (XElement row in child.Elements(OdfNamespaces.Table + "table-row")) yield return row;
        }
    }
}
