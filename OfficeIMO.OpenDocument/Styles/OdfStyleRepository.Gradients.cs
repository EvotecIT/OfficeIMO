namespace OfficeIMO.OpenDocument;

public sealed partial class OdfStyleRepository {
    /// <summary>Named native and SVG gradient definitions in common styles.</summary>
    public IReadOnlyList<OdfGradient> Gradients => GradientElements().Select(element => new OdfGradient(_document, element)).ToList();
    /// <summary>Finds a gradient, rejecting duplicate names across native and SVG gradient definitions.</summary>
    public OdfGradient? FindGradient(string name) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        XElement[] matches = GradientElements().Where(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == name).Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Duplicate gradient name '" + name + "'.");
        return matches.Length == 0 ? null : new OdfGradient(_document, matches[0]);
    }
    /// <summary>Creates a reusable native two-color gradient in common styles.</summary>
    public OdfGradient CreateGradient(string name, OdfGradientPattern pattern) {
        ValidateStyleName(name);
        if (pattern == null) throw new ArgumentNullException(nameof(pattern));
        if (FindGradient(name) != null) throw new InvalidOperationException("A gradient named '" + name + "' already exists.");
        var element = new XElement(OdfNamespaces.Draw + "gradient", new XAttribute(OdfNamespaces.Draw + "name", name));
        OdfGradient.WritePattern(element, pattern); GetContainer("styles.xml", OdfNamespaces.Office + "styles").Add(element);
        _document.MarkPartDirty("styles.xml"); return new OdfGradient(_document, element);
    }
    private IEnumerable<XElement> GradientElements() => !_document.Package.ContainsEntry("styles.xml") ? Enumerable.Empty<XElement>() :
        _document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "styles")?.Elements().Where(element =>
            element.Name == OdfNamespaces.Draw + "gradient" || element.Name == OdfNamespaces.Svg + "linearGradient" || element.Name == OdfNamespaces.Svg + "radialGradient") ?? Enumerable.Empty<XElement>();
}
