namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    private static XDocument PrepareFlatStyleScopes(XDocument content, XDocument sourceStyles) {
        var styles = new XDocument(sourceStyles);
        RenameCollisions(OdfNamespaces.Office + "font-face-decls", fonts: true);
        RenameCollisions(OdfNamespaces.Office + "automatic-styles", fonts: false);
        return styles;

        void RenameCollisions(XName containerName, bool fonts) {
            XElement? target = styles.Root?.Element(containerName);
            if (target == null) return;
            var contentDefinitions = content.Root?.Element(containerName)?.Elements()
                .Where(element => element.Attribute(OdfNamespaces.Style + "name") != null)
                .GroupBy(element => (string)element.Attribute(OdfNamespaces.Style + "name")!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal)
                ?? new Dictionary<string, XElement>(StringComparer.Ordinal);
            var names = new HashSet<string>(contentDefinitions.Keys, StringComparer.Ordinal);
            foreach (XElement element in target.Elements()) {
                string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
                if (name != null) names.Add(name);
            }
            var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XElement element in target.Elements().ToArray()) {
                string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
                if (name == null || !contentDefinitions.TryGetValue(name, out XElement? other)) continue;
                if (XNode.DeepEquals(element, other)) { element.Remove(); continue; }
                int suffix = 1;
                string replacement;
                do { replacement = name + "_flat" + suffix++.ToString(CultureInfo.InvariantCulture); } while (!names.Add(replacement));
                replacements[name] = replacement;
                element.SetAttributeValue(OdfNamespaces.Style + "name", replacement);
            }
            foreach (XElement container in styles.Root!.Elements()) {
                // Common styles cannot inherit automatic styles. Their font names
                // do, however, resolve through styles.xml's font declarations.
                if (!fonts && container.Name == OdfNamespaces.Office + "styles") continue;
                foreach (XAttribute attribute in container.DescendantsAndSelf().Attributes()) {
                    bool reference = fonts ? IsFontReference(attribute.Name) : IsStyleReference(attribute.Name);
                    if (reference && replacements.TryGetValue(attribute.Value, out string? replacement)) attribute.Value = replacement;
                }
            }
        }
    }

    private static bool IsFontReference(XName name) => name.Namespace == OdfNamespaces.Style &&
        (name.LocalName == "font-name" || name.LocalName == "font-name-asian" || name.LocalName == "font-name-complex");

    private static bool IsStyleReference(XName name) =>
        name != OdfNamespaces.Style + "parent-style-name" && name != OdfNamespaces.Style + "next-style-name" &&
        (name.LocalName.EndsWith("style-name", StringComparison.Ordinal) || name == OdfNamespaces.Style + "page-layout-name");
}
