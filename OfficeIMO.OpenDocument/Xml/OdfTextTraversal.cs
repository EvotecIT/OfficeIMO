namespace OfficeIMO.OpenDocument;

internal static class OdfTextTraversal {
    internal const int MaximumContainerDepth = 128;
    internal const int MaximumVisitedElements = 100_000;

    internal static bool IsParagraph(XElement element) => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h";

    // Only descend through inline syntax. Other elements belong to another text story
    // (notes, annotations, embedded objects) or are preserved opaque XML.
    internal static IEnumerable<XElement> Inlines(XElement parent) => Traverse(parent, IsInlineContainer);

    internal static IEnumerable<XElement> Paragraphs(XElement parent) => Traverse(parent, IsListContainer).Where(IsParagraph);

    private static bool IsInlineContainer(XElement element) => element.Name == OdfNamespaces.Text + "span" || element.Name == OdfNamespaces.Text + "a";

    private static bool IsListContainer(XElement element) => element.Name == OdfNamespaces.Text + "list" ||
        element.Name == OdfNamespaces.Text + "list-item" || element.Name == OdfNamespaces.Text + "list-header";

    /// <summary>
    /// Visits one visible text story in document order without recursive iterator calls.
    /// The nesting limit counts descended XML containers, so a list and its item count separately.
    /// </summary>
    private static IEnumerable<XElement> Traverse(XElement parent, Func<XElement, bool> isContainer) {
        var pending = new Stack<IEnumerator<XElement>>();
        pending.Push(parent.Elements().GetEnumerator());
        int visited = 0;
        try {
            while (pending.Count > 0) {
                IEnumerator<XElement> current = pending.Peek();
                if (!current.MoveNext()) {
                    pending.Pop().Dispose();
                    continue;
                }
                if (++visited > MaximumVisitedElements)
                    throw new NotSupportedException($"OpenDocument text traversal exceeds the {MaximumVisitedElements}-element safety limit.");

                XElement child = current.Current;
                bool descend = isContainer(child);
                if (descend && pending.Count > MaximumContainerDepth)
                    throw new NotSupportedException($"OpenDocument text traversal exceeds the {MaximumContainerDepth}-container nesting limit.");

                yield return child;
                if (descend) pending.Push(child.Elements().GetEnumerator());
            }
        } finally {
            while (pending.Count > 0) pending.Pop().Dispose();
        }
    }
}
