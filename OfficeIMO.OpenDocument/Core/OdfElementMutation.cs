namespace OfficeIMO.OpenDocument;

// Resolves a previously read child against a logical cell's materialized XML.
// Reading wrappers remains sparse; only an edit invokes the resolver.
internal static class OdfElementMutation {
    internal static Func<XElement>? ForDescendant(XElement root, XElement child, Func<XElement>? materialize) {
        if (materialize == null) return null;
        var indices = new List<int>();
        for (XElement current = child; current != root; current = current.Parent
            ?? throw new InvalidOperationException("The editable element is no longer in its original container.")) {
            indices.Add(current.ElementsBeforeSelf().Count());
        }
        indices.Reverse();
        return () => {
            XElement current = materialize();
            foreach (int index in indices) current = current.Elements().ElementAt(index);
            return current;
        };
    }
}
