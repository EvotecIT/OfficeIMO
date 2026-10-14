namespace OfficeIMO.OpenDocument;

/// <summary>Inserts optional style containers in content/styles part order without moving existing XML.</summary>
internal static class OdfXmlContainers {
    internal static XElement Ensure(XElement root, XName name) {
        XElement? existing = root.Element(name);
        if (existing != null) return existing;
        var container = new XElement(name);
        int rank = Rank(name);
        XElement? following = root.Elements().FirstOrDefault(element => Rank(element.Name) > rank);
        if (following == null) root.Add(container); else following.AddBeforeSelf(container);
        return container;
    }

    private static int Rank(XName name) => name.Namespace != OdfNamespaces.Office ? 6 : name.LocalName switch {
        "scripts" => 0, "font-face-decls" => 1, "styles" => 2,
        "automatic-styles" => 3, "master-styles" => 4, "body" => 5, _ => 6
    };
}
