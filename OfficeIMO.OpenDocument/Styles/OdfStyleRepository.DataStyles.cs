namespace OfficeIMO.OpenDocument;

public sealed partial class OdfStyleRepository {
    /// <summary>Finds a data style in the requested part's automatic styles, then common styles. Duplicate names are rejected.</summary>
    public OdfDataStyle? FindDataStyle(string name, string partPath = "content.xml") {
        if (partPath is not ("content.xml" or "styles.xml")) throw new ArgumentException("Use an ODF content or styles part.", nameof(partPath));
        if (string.IsNullOrWhiteSpace(name)) return null;
        foreach (var location in new[] { (partPath, OdfNamespaces.Office + "automatic-styles"), ("styles.xml", OdfNamespaces.Office + "styles") }) {
            OdfDataStyle? style = FindDataStyleInContainer(name, location.Item1, location.Item2);
            if (style != null) return style;
        }
        return null;
    }

    // Conditional maps require a common target, even when an automatic style
    // with the same name is visible to direct field bindings (ODF 19.466).
    internal OdfDataStyle? FindCommonDataStyle(string name) =>
        FindDataStyleInContainer(name, "styles.xml", OdfNamespaces.Office + "styles");

    private OdfDataStyle? FindDataStyleInContainer(string name, string partPath, XName container) {
        if (!_document.Package.ContainsEntry(partPath)) return null;
        XElement[] matches = (_document.GetXml(partPath).Root?.Element(container)?.Elements() ?? Enumerable.Empty<XElement>())
            .Where(element => element.Name.Namespace == OdfNamespaces.Number && (string?)element.Attribute(OdfNamespaces.Style + "name") == name)
            .Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Duplicate data style '" + name + "' in " + partPath + ".");
        return matches.Length == 1 ? new OdfDataStyle(_document, matches[0]) : null;
    }

    /// <summary>Creates a common Gregorian ISO year-month-day style usable by any ODF document.</summary>
    public OdfDataStyle CreateDateStyle(string name) => CreateCommonDataStyle(name, "date-style",
        DatePart("year"), Literal("-"), DatePart("month"), Literal("-"), DatePart("day"));
    /// <summary>Creates a common 24-hour clock style with hours, minutes and seconds.</summary>
    public OdfDataStyle CreateTimeStyle(string name) => CreateCommonDataStyle(name, "time-style",
        TimePart("hours"), Literal(":"), TimePart("minutes"), Literal(":"), TimePart("seconds"));

    private OdfDataStyle CreateCommonDataStyle(string name, string kind, params XElement[] components) {
        ValidateStyleName(name);
        XElement container = GetContainer("styles.xml", OdfNamespaces.Office + "styles");
        if (container.Elements().Any(element => element.Name.Namespace == OdfNamespaces.Number &&
            (string?)element.Attribute(OdfNamespaces.Style + "name") == name)) throw new InvalidOperationException("A common data style named '" + name + "' already exists.");
        var definition = new XElement(OdfNamespaces.Number + kind, new XAttribute(OdfNamespaces.Style + "name", name), components);
        container.Add(definition); _document.MarkPartDirty("styles.xml");
        return new OdfDataStyle(_document, definition);
    }
    private static XElement DatePart(string name) => new(OdfNamespaces.Number + name,
        new XAttribute(OdfNamespaces.Number + "style", "long"), new XAttribute(OdfNamespaces.Number + "calendar", "gregorian"));
    private static XElement TimePart(string name) => new(OdfNamespaces.Number + name, new XAttribute(OdfNamespaces.Number + "style", "long"));
    private static XElement Literal(string text) => new(OdfNamespaces.Number + "text", text);
}
