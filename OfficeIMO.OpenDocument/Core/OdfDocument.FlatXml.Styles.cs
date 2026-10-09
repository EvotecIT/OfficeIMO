namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    private static XElement MergeFlatAutomaticStyles(XElement? content, XElement? styles) {
        XElement merged = MergeContainers(OdfNamespaces.Office + "automatic-styles", content, styles);
        // Keep merged style definitions together before page layouts for interoperable validation.
        XElement[] layouts = merged.Elements(OdfNamespaces.Style + "page-layout").ToArray();
        foreach (XElement layout in layouts) { layout.Remove(); merged.Add(layout); }
        return merged;
    }

    private static XDocument PrepareFlatStyleScopes(XDocument content, XDocument sourceStyles) {
        var styles = new XDocument(sourceStyles);
        RenameCommonCollisions();
        RenameCollisions(OdfNamespaces.Office + "font-face-decls", fonts: true);
        RenameCollisions(OdfNamespaces.Office + "automatic-styles", fonts: false);
        return styles;

        // Flattening merges two automatic scopes. Neither scope may then shadow
        // a common definition that was visible only from the other source part.
        void RenameCommonCollisions() {
            var commonKeys = new HashSet<string>((styles.Root?.Element(OdfNamespaces.Office + "styles")?.Elements() ?? Enumerable.Empty<XElement>())
                .Where(element => element.Attribute(OdfNamespaces.Style + "name") != null).Select(element => DefinitionKey(element, false)), StringComparer.Ordinal);
            var names = new HashSet<string>(new[] { content, styles }.SelectMany(part => part.Descendants().Attributes(OdfNamespaces.Style + "name"))
                .Select(attribute => attribute.Value), StringComparer.Ordinal);
            foreach (XDocument part in new[] { content, styles }) {
                var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (XElement definition in part.Root?.Element(OdfNamespaces.Office + "automatic-styles")?.Elements() ?? Enumerable.Empty<XElement>()) {
                    string key = DefinitionKey(definition, false);
                    if (!commonKeys.Contains(key)) continue;
                    string name = (string)definition.Attribute(OdfNamespaces.Style + "name")!;
                    int suffix = 1; string replacement;
                    do { replacement = name + "_flatCommon" + suffix++.ToString(CultureInfo.InvariantCulture); } while (!names.Add(replacement));
                    definition.SetAttributeValue(OdfNamespaces.Style + "name", replacement); replacements[key] = replacement;
                }
                RewriteReferences(part, replacements, false);
            }
        }

        void RenameCollisions(XName containerName, bool fonts) {
            XElement? target = styles.Root?.Element(containerName);
            if (target == null) return;
            var contentDefinitions = content.Root?.Element(containerName)?.Elements()
                .Where(element => element.Attribute(OdfNamespaces.Style + "name") != null)
                .GroupBy(element => DefinitionKey(element, fonts), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal)
                ?? new Dictionary<string, XElement>(StringComparer.Ordinal);
            var names = new HashSet<string>(contentDefinitions.Values.Select(element =>
                (string)element.Attribute(OdfNamespaces.Style + "name")!), StringComparer.Ordinal);
            foreach (XAttribute commonName in styles.Root!.Element(OdfNamespaces.Office + "styles")?.DescendantsAndSelf().Attributes(OdfNamespaces.Style + "name") ?? Enumerable.Empty<XAttribute>())
                names.Add(commonName.Value);
            foreach (XElement element in target.Elements()) {
                string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
                if (name != null) names.Add(name);
            }
            var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XElement element in target.Elements().ToArray()) {
                string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
                string key = DefinitionKey(element, fonts);
                if (name == null || !contentDefinitions.TryGetValue(key, out XElement? other)) continue;
                // Even identical definitions can reference different part-local
                // number/list/font styles. Keep their dependency graphs separate.
                int suffix = 1;
                string replacement;
                do { replacement = name + "_flat" + suffix++.ToString(CultureInfo.InvariantCulture); } while (!names.Add(replacement));
                replacements[key] = replacement;
                element.SetAttributeValue(OdfNamespaces.Style + "name", replacement);
            }
            RewriteReferences(styles, replacements, fonts);
        }
        void RewriteReferences(XDocument part, Dictionary<string, string> replacements, bool fonts) {
            foreach (XElement container in part.Root!.Elements()) {
                foreach (XAttribute attribute in container.DescendantsAndSelf().Attributes()) {
                    string? kind = fonts ? (IsFontReference(attribute.Name) ? "font" : null) : ReferenceKind(attribute);
                    if (kind == null) continue;
                    if (attribute.Name.LocalName == "class-names") {
                        string[] values = attribute.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                        attribute.Value = string.Join(" ", values.Select(value =>
                            replacements.TryGetValue(kind + "\0" + value, out string? mapped) ? mapped : value));
                    } else if (replacements.TryGetValue(kind + "\0" + attribute.Value, out string? replacement)) {
                        attribute.Value = replacement;
                    }
                }
            }
        }
    }

    private static bool IsFontReference(XName name) => name.Namespace == OdfNamespaces.Style &&
        (name.LocalName == "font-name" || name.LocalName == "font-name-asian" || name.LocalName == "font-name-complex");

    private static string DefinitionKey(XElement element, bool fonts) =>
        (fonts ? "font" : DefinitionKind(element)) + "\0" + (string?)element.Attribute(OdfNamespaces.Style + "name");

    private static string DefinitionKind(XElement element) {
        if (element.Name == OdfNamespaces.Style + "style") return "family:" + (string?)element.Attribute(OdfNamespaces.Style + "family");
        if (element.Name.Namespace == OdfNamespaces.Number) return "data";
        if (element.Name == OdfNamespaces.Text + "list-style") return "list";
        return element.Name.ToString();
    }

    // ODF permits the same lexical name in different style families. Resolve
    // references by their declared role, rather than rewriting every matching string.
    private static string? ReferenceKind(XAttribute attribute) {
        XName name = attribute.Name;
        XElement owner = attribute.Parent!;
        string local = name.LocalName;
        if (name.Namespace == OdfNamespaces.Style) {
            if (local == "leader-text-style") {
                XElement? definition = owner.Ancestors().FirstOrDefault(element =>
                    element.Name == OdfNamespaces.Style + "style" || element.Name == OdfNamespaces.Style + "default-style");
                // Automatic styles can bind automatic text styles, but common/default
                // styles retain common-only bindings when automatic scopes are merged.
                return definition?.Parent?.Name == OdfNamespaces.Office + "styles" ? null : "family:text";
            }
            if (local == "data-style-name" || local == "percentage-data-style-name") return "data";
            if (local == "list-style-name") return "list";
            if (local == "page-layout-name") return (OdfNamespaces.Style + "page-layout").ToString();
            if (local == "style-name" && owner.Name == OdfNamespaces.Style + "drop-cap") return "family:text";
            // Parent, next and conditional style-map targets refer to common styles.
            return null;
        }
        if (name.Namespace == OdfNamespaces.Presentation) {
            if (local == "presentation-page-layout-name") return (OdfNamespaces.Style + "presentation-page-layout").ToString();
            if (local == "style-name" || local == "class-names") return "family:presentation";
        }
        if (name.Namespace == OdfNamespaces.Draw) {
            if (local == "text-style-name") return "family:paragraph";
            if (local == "style-name" || local == "class-names") return
                owner.Name == OdfNamespaces.Draw + "page" || owner.Name == OdfNamespaces.Style + "master-page" ||
                owner.Name == OdfNamespaces.Style + "handout-master" || owner.Name == OdfNamespaces.Presentation + "notes"
                    ? "family:drawing-page" : "family:graphic";
        }
        if (name.Namespace == OdfNamespaces.Chart && local == "style-name") return "family:chart";
        if (name.Namespace == OdfNamespaces.Table) {
            if (local == "default-cell-style-name") return "family:table-cell";
            if (local == "paragraph-style-name") return "family:paragraph";
            if (local == "style-name") return "family:" + (owner.Name.LocalName == "covered-table-cell" ? "table-cell" : owner.Name.LocalName);
        }
        if (name.Namespace == OdfNamespaces.Text) {
            if (local == "list-style-name") return "list";
            if (local == "paragraph-style-name" || local == "cond-style-name" || local == "default-style-name") return "family:paragraph";
            if (local == "visited-style-name" || local == "citation-style-name" || local == "citation-body-style-name" || local == "main-entry-style-name") return "family:text";
            if (local == "style-name" || local == "class-names") {
                string element = owner.Name.LocalName;
                if (element == "list" || element == "numbered-paragraph") return "list";
                if (element == "section" || element.EndsWith("index", StringComparison.Ordinal)) return "family:section";
                if (element == "ruby") return "family:ruby";
                if (element == "p" || element == "h" || element == "index-title-template" || element == "index-source-style" ||
                    element.EndsWith("entry-template", StringComparison.Ordinal)) return "family:paragraph";
                return "family:text";
            }
        }
        return null;
    }
}
