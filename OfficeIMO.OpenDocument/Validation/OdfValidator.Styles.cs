namespace OfficeIMO.OpenDocument;

internal static partial class OdfValidator {
    private static void ValidateStyles(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        var definitions = new List<(string Part, string Family, string Name, string? Parent)>();
        var knownNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (string partPath in new[] { "styles.xml", "content.xml" }) {
            if (!package.ContainsEntry(partPath)) continue;
            XDocument document = package.GetXml(partPath);
            foreach (XElement element in document.Descendants()) {
                string? declaredName = (string?)element.Attribute(OdfNamespaces.Style + "name") ??
                    (string?)element.Attribute(OdfNamespaces.Draw + "name");
                if (!string.IsNullOrWhiteSpace(declaredName)) knownNames.Add(declaredName!);
                if (element.Name != OdfNamespaces.Style + "style") continue;
                string? name = (string?)element.Attribute(OdfNamespaces.Style + "name");
                if (string.IsNullOrWhiteSpace(name)) continue;
                string family = (string?)element.Attribute(OdfNamespaces.Style + "family") ?? string.Empty;
                definitions.Add((partPath, family, name!, (string?)element.Attribute(OdfNamespaces.Style + "parent-style-name")));
            }
        }

        foreach (IGrouping<string, (string Part, string Family, string Name, string? Parent)> duplicate in definitions
            .GroupBy(item => item.Family + "\0" + item.Name, StringComparer.Ordinal).Where(group => group.Count() > 1)) {
            diagnostics.Add(new OdfDiagnostic("ODF201", OdfDiagnosticSeverity.Error,
                $"Style '{duplicate.First().Name}' in family '{duplicate.First().Family}' is declared {duplicate.Count()} times.", duplicate.First().Part));
        }

        var byKey = definitions.GroupBy(item => item.Family + "\0" + item.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        foreach (var definition in definitions) {
            if (string.IsNullOrWhiteSpace(definition.Parent)) continue;
            string parentKey = definition.Family + "\0" + definition.Parent;
            if (!byKey.ContainsKey(parentKey)) {
                diagnostics.Add(new OdfDiagnostic("ODF202", OdfDiagnosticSeverity.Error,
                    $"Style '{definition.Name}' references missing parent style '{definition.Parent}'.", definition.Part));
            }
        }

        var completed = new HashSet<string>(StringComparer.Ordinal);
        foreach (string start in byKey.Keys) {
            if (completed.Contains(start)) continue;
            var path = new HashSet<string>(StringComparer.Ordinal);
            string? current = start;
            while (current != null && byKey.TryGetValue(current, out var definition)) {
                if (!path.Add(current)) {
                    diagnostics.Add(new OdfDiagnostic("ODF203", OdfDiagnosticSeverity.Error,
                        $"Style parent cycle detected at '{definition.Name}'.", definition.Part));
                    break;
                }
                if (string.IsNullOrWhiteSpace(definition.Parent)) break;
                current = definition.Family + "\0" + definition.Parent;
            }
            foreach (string key in path) completed.Add(key);
        }

        foreach (string partPath in new[] { "styles.xml", "content.xml" }) {
            if (!package.ContainsEntry(partPath)) continue;
            foreach (XAttribute reference in package.GetXml(partPath).Descendants().Attributes().Where(IsStyleReference)) {
                if (!string.IsNullOrWhiteSpace(reference.Value) && !knownNames.Contains(reference.Value)) {
                    diagnostics.Add(new OdfDiagnostic("ODF200", OdfDiagnosticSeverity.Error,
                        $"Style reference '{reference.Value}' from '{reference.Name.LocalName}' does not resolve.", partPath));
                }
            }
        }
    }

    private static bool IsStyleReference(XAttribute attribute) {
        switch (attribute.Name.LocalName) {
            case "style-name":
            case "parent-style-name":
            case "next-style-name":
            case "list-style-name":
            case "data-style-name":
            case "default-cell-style-name":
            case "paragraph-style-name":
                return true;
            default:
                return false;
        }
    }
}
