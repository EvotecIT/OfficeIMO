namespace OfficeIMO.OpenDocument;

public sealed partial class OdfStyleRepository {
    /// <summary>Named stroke-dash definitions in common styles.</summary>
    public IReadOnlyList<OdfStrokeDash> StrokeDashes => StrokeDashElements().Select(element => new OdfStrokeDash(_document, element)).ToList();
    /// <summary>Finds a named dash definition, rejecting ambiguous duplicate names.</summary>
    public OdfStrokeDash? FindStrokeDash(string name) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        XElement[] matches = StrokeDashElements().Where(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == name).Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Duplicate stroke-dash name '" + name + "'.");
        return matches.Length == 0 ? null : new OdfStrokeDash(_document, matches[0]);
    }
    /// <summary>Creates a named dash pattern shared by shapes that reference it.</summary>
    public OdfStrokeDash CreateStrokeDash(string name, OdfStrokeDashPattern pattern) {
        ValidateStyleName(name);
        if (pattern == null) throw new ArgumentNullException(nameof(pattern));
        if (FindStrokeDash(name) != null) throw new InvalidOperationException("A stroke-dash named '" + name + "' already exists.");
        var element = new XElement(OdfNamespaces.Draw + "stroke-dash", new XAttribute(OdfNamespaces.Draw + "name", name));
        OdfStrokeDash.WritePattern(element, pattern);
        GetContainer("styles.xml", OdfNamespaces.Office + "styles").Add(element);
        _document.MarkPartDirty("styles.xml");
        return new OdfStrokeDash(_document, element);
    }
    private IEnumerable<XElement> StrokeDashElements() => !_document.Package.ContainsEntry("styles.xml") ? Enumerable.Empty<XElement>() :
        _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "styles")?.Elements(OdfNamespaces.Draw + "stroke-dash") ?? Enumerable.Empty<XElement>();
}
