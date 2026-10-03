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
        var pending = new Stack<XElement>();
        pending.Push(child);
        while (pending.Count > 0) {
            XElement current = pending.Pop();
            if (current.Name == OdfNamespaces.Table + "table-row") {
                yield return current;
            } else if (current.Name == OdfNamespaces.Table + "table-header-rows"
                || current.Name == OdfNamespaces.Table + "table-rows"
                || current.Name == OdfNamespaces.Table + "table-row-group") {
                XElement[] children = current.Elements().ToArray();
                for (int index = children.Length - 1; index >= 0; index--) pending.Push(children[index]);
            }
        }
    }
}
