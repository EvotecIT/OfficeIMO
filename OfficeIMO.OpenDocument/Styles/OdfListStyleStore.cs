namespace OfficeIMO.OpenDocument;

internal static class OdfListStyleStore {
    internal static string Create(OdfDocument document, bool ordered, string partPath = "content.xml", int listLevel = 1) {
        if (listLevel < 1 || listLevel > 10) throw new ArgumentOutOfRangeException(nameof(listLevel));
        XDocument xml = document.GetXml(partPath);
        XElement root = xml.Root ?? throw new InvalidDataException($"OpenDocument part '{partPath}' has no root element.");
        XElement styles = OdfXmlContainers.Ensure(root, OdfNamespaces.Office + "automatic-styles");
        // A new automatic definition must not shadow a referenced common list style.
        // Reserve both parts as flat serialization combines their automatic definitions.
        var existingNames = new HashSet<string>(new[] { "content.xml", "styles.xml" }
            .Where(document.Package.ContainsEntry).SelectMany(part => document.GetXml(part).Root!.Elements()
                .Where(container => container.Name == OdfNamespaces.Office + "automatic-styles" || container.Name == OdfNamespaces.Office + "styles")
                .SelectMany(container => container.Elements(OdfNamespaces.Text + "list-style")))
            .Select(element => (string?)element.Attribute(OdfNamespaces.Style + "name"))
            .Where(value => !string.IsNullOrEmpty(value))!, StringComparer.Ordinal);
        int index = 1; string name;
        do { name = "ofList" + index++.ToString("D4", CultureInfo.InvariantCulture); } while (existingNames.Contains(name));
        XElement level = ordered
            ? new XElement(OdfNamespaces.Text + "list-level-style-number",
                new XAttribute(OdfNamespaces.Text + "level", listLevel), new XAttribute(OdfNamespaces.Style + "num-format", "1"), new XAttribute(OdfNamespaces.Style + "num-suffix", "."))
            : new XElement(OdfNamespaces.Text + "list-level-style-bullet",
                new XAttribute(OdfNamespaces.Text + "level", listLevel), new XAttribute(OdfNamespaces.Text + "bullet-char", "•"));
        level.Add(new XElement(OdfNamespaces.Style + "list-level-properties",
            new XAttribute(OdfNamespaces.Text + "space-before", (listLevel - 1) * 18 + "pt"),
            new XAttribute(OdfNamespaces.Text + "min-label-width", "18pt"),
            new XAttribute(OdfNamespaces.Text + "min-label-distance", "3pt")));
        styles.Add(new XElement(OdfNamespaces.Text + "list-style", new XAttribute(OdfNamespaces.Style + "name", name), level));
        document.MarkPartDirty(partPath);
        return name;
    }

    internal static bool IsOrdered(OdfDocument document, string? styleName, string partPath = "content.xml", int level = 1) {
        if (string.IsNullOrWhiteSpace(styleName)) return false;
        XElement? style = Find(document, partPath, OdfNamespaces.Office + "automatic-styles", styleName!);
        if (style == null && !string.Equals(partPath, "styles.xml", StringComparison.Ordinal) && document.Package.ContainsEntry("styles.xml")) {
            style = Find(document, "styles.xml", OdfNamespaces.Office + "automatic-styles", styleName!) ??
                Find(document, "styles.xml", OdfNamespaces.Office + "styles", styleName!);
        } else if (style == null && string.Equals(partPath, "styles.xml", StringComparison.Ordinal)) {
            style = Find(document, "styles.xml", OdfNamespaces.Office + "styles", styleName!);
        }
        XElement? levelStyle = style?.Elements().FirstOrDefault(element =>
            int.TryParse((string?)element.Attribute(OdfNamespaces.Text + "level"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int definedLevel) && definedLevel == level);
        return levelStyle != null ? levelStyle.Name == OdfNamespaces.Text + "list-level-style-number" :
            style?.Elements().Any(element => element.Name == OdfNamespaces.Text + "list-level-style-number") == true;
    }

    private static XElement? Find(OdfDocument document, string partPath, XName containerName, string styleName) {
        XElement? container = document.GetXml(partPath).Root?.Element(containerName);
        return container?.Elements(OdfNamespaces.Text + "list-style")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Style + "name"), styleName, StringComparison.Ordinal));
    }
}
