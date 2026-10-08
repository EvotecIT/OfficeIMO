namespace OfficeIMO.OpenDocument;

public sealed partial class OdfStyleRepository {
    /// <summary>Creates a unique common binding for an embedded resource without changing existing bindings.</summary>
    internal string CreateFillImage(string path) {
        XElement container = GetContainer("styles.xml", OdfNamespaces.Office + "styles");
        var names = new HashSet<string>(container.Elements(OdfNamespaces.Draw + "fill-image")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") ?? ""), StringComparer.Ordinal);
        int index = 1; string name;
        do { name = "ofBgImage" + index++.ToString(CultureInfo.InvariantCulture); } while (names.Contains(name));
        container.Add(new XElement(OdfNamespaces.Draw + "fill-image", new XAttribute(OdfNamespaces.Draw + "name", name),
            new XAttribute(OdfNamespaces.XLink + "href", path), new XAttribute(OdfNamespaces.XLink + "type", "simple"),
            new XAttribute(OdfNamespaces.XLink + "show", "embed"), new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
        _document.MarkPartDirty("styles.xml"); return name;
    }

    internal byte[] ReadFillImage(string name) {
        XElement[] matches = _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "styles")?
            .Elements(OdfNamespaces.Draw + "fill-image").Where(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == name)
            .Take(2).ToArray() ?? Array.Empty<XElement>();
        if (matches.Length != 1) throw new InvalidDataException("A fill image must resolve to one common definition: " + name + ".");
        string? href = (string?)matches[0].Attribute(OdfNamespaces.XLink + "href");
        if (string.IsNullOrWhiteSpace(href)) throw new NotSupportedException("Fill image has no embedded package resource.");
        string path = OdfPackagePath.NormalizeHref(href!);
        if (!_document.Package.ContainsEntry(path)) throw new NotSupportedException("Linked and unavailable fill images are preserved without fetching external content.");
        return _document.GetPackageEntryBytes(path);
    }
}
