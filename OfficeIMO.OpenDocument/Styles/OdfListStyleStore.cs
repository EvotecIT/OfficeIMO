namespace OfficeIMO.OpenDocument;

internal static class OdfListStyleStore {
    internal static string Create(OdfDocument document, bool ordered, string partPath = "content.xml") {
        XDocument xml = document.GetXml(partPath);
        XElement root = xml.Root ?? throw new InvalidDataException($"OpenDocument part '{partPath}' has no root element.");
        XElement styles = root.Element(OdfNamespaces.Office + "automatic-styles") ?? throw new InvalidDataException($"OpenDocument part '{partPath}' has no automatic styles.");
        var existingNames = new HashSet<string>(styles.Elements(OdfNamespaces.Text + "list-style")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Style + "name"))
            .Where(value => !string.IsNullOrEmpty(value))!, StringComparer.Ordinal);
        int index = 1; string name;
        do { name = "ofList" + index++.ToString("D4", CultureInfo.InvariantCulture); } while (existingNames.Contains(name));
        XElement level = ordered
            ? new XElement(OdfNamespaces.Text + "list-level-style-number",
                new XAttribute(OdfNamespaces.Text + "level", 1), new XAttribute(OdfNamespaces.Style + "num-format", "1"))
            : new XElement(OdfNamespaces.Text + "list-level-style-bullet",
                new XAttribute(OdfNamespaces.Text + "level", 1), new XAttribute(OdfNamespaces.Text + "bullet-char", "•"));
        styles.Add(new XElement(OdfNamespaces.Text + "list-style", new XAttribute(OdfNamespaces.Style + "name", name), level));
        document.MarkPartDirty(partPath);
        return name;
    }

    internal static bool IsOrdered(OdfDocument document, string? styleName, string partPath = "content.xml") {
        if (string.IsNullOrWhiteSpace(styleName)) return false;
        XElement? style = Find(document, partPath, OdfNamespaces.Office + "automatic-styles", styleName!);
        if (style == null && !string.Equals(partPath, "styles.xml", StringComparison.Ordinal) && document.Package.ContainsEntry("styles.xml")) {
            style = Find(document, "styles.xml", OdfNamespaces.Office + "automatic-styles", styleName!) ??
                Find(document, "styles.xml", OdfNamespaces.Office + "styles", styleName!);
        } else if (style == null && string.Equals(partPath, "styles.xml", StringComparison.Ordinal)) {
            style = Find(document, "styles.xml", OdfNamespaces.Office + "styles", styleName!);
        }
        return style?.Elements().Any(element => element.Name == OdfNamespaces.Text + "list-level-style-number") == true;
    }

    private static XElement? Find(OdfDocument document, string partPath, XName containerName, string styleName) {
        XElement? container = document.GetXml(partPath).Root?.Element(containerName);
        return container?.Elements(OdfNamespaces.Text + "list-style")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Style + "name"), styleName, StringComparison.Ordinal));
    }
}
