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
}
