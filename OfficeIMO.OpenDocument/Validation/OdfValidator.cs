namespace OfficeIMO.OpenDocument;

internal static class OdfValidator {
    internal static OdfValidationResult Validate(OdfPackage package) {
        var diagnostics = new List<OdfDiagnostic>();
        RequireEntry(package, "mimetype", diagnostics);
        RequireEntry(package, "content.xml", diagnostics);
        RequireEntry(package, "META-INF/manifest.xml", diagnostics);

        ValidateXmlRoot(package, "content.xml", OdfNamespaces.Office + "document-content", diagnostics);
        if (package.ContainsEntry("styles.xml")) ValidateXmlRoot(package, "styles.xml", OdfNamespaces.Office + "document-styles", diagnostics);
        if (package.ContainsEntry("meta.xml")) ValidateXmlRoot(package, "meta.xml", OdfNamespaces.Office + "document-meta", diagnostics);
        if (package.ContainsEntry("settings.xml")) ValidateXmlRoot(package, "settings.xml", OdfNamespaces.Office + "document-settings", diagnostics);

        if (package.ContainsEntry("content.xml")) {
            XDocument content = package.GetXml("content.xml");
            XElement? body = content.Root?.Element(OdfNamespaces.Office + "body");
            XName expectedBody;
            switch (package.Kind) {
                case OdfDocumentKind.Text: expectedBody = OdfNamespaces.Office + "text"; break;
                case OdfDocumentKind.Spreadsheet: expectedBody = OdfNamespaces.Office + "spreadsheet"; break;
                default: expectedBody = OdfNamespaces.Office + "presentation"; break;
            }
            if (body?.Element(expectedBody) == null) {
                diagnostics.Add(new OdfDiagnostic("ODF102", OdfDiagnosticSeverity.Error,
                    $"content.xml does not contain the expected '{expectedBody}' body.", "content.xml"));
            }
        }

        ValidateManifest(package, diagnostics);
        ValidateVersionConsistency(package, diagnostics);
        return new OdfValidationResult(diagnostics);
    }

    private static void RequireEntry(OdfPackage package, string name, List<OdfDiagnostic> diagnostics) {
        if (!package.ContainsEntry(name)) {
            diagnostics.Add(new OdfDiagnostic("ODF100", OdfDiagnosticSeverity.Error, $"Required package entry '{name}' is missing.", name));
        }
    }

    private static void ValidateXmlRoot(OdfPackage package, string partPath, XName expectedName, List<OdfDiagnostic> diagnostics) {
        try {
            XDocument document = package.GetXml(partPath);
            if (document.Root?.Name != expectedName) {
                diagnostics.Add(new OdfDiagnostic("ODF101", OdfDiagnosticSeverity.Error,
                    $"Part '{partPath}' has root '{document.Root?.Name}' instead of '{expectedName}'.", partPath));
            }
        } catch (Exception ex) {
            diagnostics.Add(new OdfDiagnostic("ODF101", OdfDiagnosticSeverity.Error, ex.Message, partPath));
        }
    }

    private static void ValidateManifest(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        if (!package.ContainsEntry("META-INF/manifest.xml")) return;
        XDocument manifest = package.GetXml("META-INF/manifest.xml");
        XElement? root = manifest.Root;
        if (root == null) return;
        Dictionary<string, int> listed = root.Elements(OdfNamespaces.Manifest + "file-entry")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path"))
            .Where(path => !string.IsNullOrEmpty(path))
            .GroupBy(path => path!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        foreach (OdfPackageEntry entry in package.Entries) {
            if (entry.Name == "mimetype" || entry.Name.StartsWith("META-INF/", StringComparison.Ordinal)) continue;
            if (!listed.TryGetValue(entry.Name, out int count)) {
                diagnostics.Add(new OdfDiagnostic("ODF103", OdfDiagnosticSeverity.Error,
                    $"Manifest does not list package entry '{entry.Name}'.", "META-INF/manifest.xml"));
            } else if (count != 1) {
                diagnostics.Add(new OdfDiagnostic("ODF104", OdfDiagnosticSeverity.Error,
                    $"Manifest lists package entry '{entry.Name}' {count} times.", "META-INF/manifest.xml"));
            }
        }
    }

    private static void ValidateVersionConsistency(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        string expected = package.Version.ToToken();
        foreach (string partPath in new[] { "content.xml", "styles.xml", "meta.xml", "settings.xml" }) {
            if (!package.ContainsEntry(partPath)) continue;
            string? actual = (string?)package.GetXml(partPath).Root?.Attribute(OdfNamespaces.Office + "version");
            if (!string.Equals(actual, expected, StringComparison.Ordinal)) {
                diagnostics.Add(new OdfDiagnostic("ODF105", OdfDiagnosticSeverity.Warning,
                    $"Part version '{actual ?? "<missing>"}' does not match package version '{expected}'.", partPath));
            }
        }
    }
}
