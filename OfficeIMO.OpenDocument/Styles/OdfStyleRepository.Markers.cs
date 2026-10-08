namespace OfficeIMO.OpenDocument;

public sealed partial class OdfStyleRepository {
    /// <summary>Named marker definitions in common styles.</summary>
    public IReadOnlyList<OdfMarker> Markers => MarkerElements().Select(element => new OdfMarker(_document, element)).ToList();
    /// <summary>Finds a marker, rejecting ambiguous duplicate names.</summary>
    public OdfMarker? FindMarker(string name) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        XElement[] matches = MarkerElements().Where(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == name).Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Duplicate marker name '" + name + "'.");
        return matches.Length == 0 ? null : new OdfMarker(_document, matches[0]);
    }
    /// <summary>Creates reusable native marker geometry in common styles.</summary>
    public OdfMarker CreateMarker(string name, OdfMarkerGeometry geometry) {
        ValidateStyleName(name);
        if (geometry == null) throw new ArgumentNullException(nameof(geometry));
        if (FindMarker(name) != null) throw new InvalidOperationException("A marker named '" + name + "' already exists.");
        var element = new XElement(OdfNamespaces.Draw + "marker", new XAttribute(OdfNamespaces.Draw + "name", name));
        OdfMarker.WriteGeometry(element, geometry); GetContainer("styles.xml", OdfNamespaces.Office + "styles").Add(element);
        _document.MarkPartDirty("styles.xml"); return new OdfMarker(_document, element);
    }
    private IEnumerable<XElement> MarkerElements() => !_document.Package.ContainsEntry("styles.xml") ? Enumerable.Empty<XElement>() :
        _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "styles")?.Elements(OdfNamespaces.Draw + "marker") ?? Enumerable.Empty<XElement>();
}
