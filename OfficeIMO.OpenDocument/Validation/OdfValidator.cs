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
        if (package.Kind == OdfDocumentKind.Spreadsheet) ValidateSpreadsheet(package, diagnostics);
        if (package.Kind == OdfDocumentKind.Presentation) ValidatePresentation(package, diagnostics);
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

    private static void ValidateSpreadsheet(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        XDocument content = package.GetXml("content.xml");
        foreach (XElement element in content.Descendants()) {
            foreach (XName repeatName in new[] {
                OdfNamespaces.Table + "number-rows-repeated", OdfNamespaces.Table + "number-columns-repeated"
            }) {
                XAttribute? attribute = element.Attribute(repeatName);
                if (attribute != null && (!long.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out long count) || count < 1)) {
                    diagnostics.Add(new OdfDiagnostic("ODS100", OdfDiagnosticSeverity.Error,
                        $"Invalid spreadsheet repeat count '{attribute.Value}'.", "content.xml"));
                }
            }
            if (element.Name != OdfNamespaces.Table + "table-cell") continue;
            string? valueType = (string?)element.Attribute(OdfNamespaces.Office + "value-type");
            XName? required = null;
            switch (valueType) {
                case "float": case "percentage": case "currency": required = OdfNamespaces.Office + "value"; break;
                case "boolean": required = OdfNamespaces.Office + "boolean-value"; break;
                case "date": required = OdfNamespaces.Office + "date-value"; break;
                case "time": required = OdfNamespaces.Office + "time-value"; break;
            }
            if (required != null && element.Attribute(required) == null) {
                diagnostics.Add(new OdfDiagnostic("ODS101", OdfDiagnosticSeverity.Error,
                    $"Spreadsheet cell type '{valueType}' is missing '{required.LocalName}'.", "content.xml"));
            }
        }
    }

    private static void ValidatePresentation(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        XDocument content = package.GetXml("content.xml");
        XDocument? styles = package.ContainsEntry("styles.xml") ? package.GetXml("styles.xml") : null;
        var masters = new HashSet<string>(styles?.Descendants(OdfNamespaces.Style + "master-page")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Style + "name"))
            .Where(value => !string.IsNullOrEmpty(value)).Select(value => value!) ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
        var slideNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement slide in content.Descendants(OdfNamespaces.Draw + "page")) {
            string? name = (string?)slide.Attribute(OdfNamespaces.Draw + "name");
            if (string.IsNullOrEmpty(name) || !slideNames.Add(name!)) {
                diagnostics.Add(new OdfDiagnostic("ODP100", OdfDiagnosticSeverity.Error,
                    "Presentation slide names must be present and unique.", "content.xml"));
            }
            string? master = (string?)slide.Attribute(OdfNamespaces.Draw + "master-page-name");
            if (!string.IsNullOrEmpty(master) && !masters.Contains(master!)) {
                diagnostics.Add(new OdfDiagnostic("ODP101", OdfDiagnosticSeverity.Error,
                    $"Presentation slide references missing master page '{master}'.", "content.xml"));
            }
        }
        foreach (XElement image in content.Descendants(OdfNamespaces.Draw + "image")) {
            string? href = (string?)image.Attribute(OdfNamespaces.XLink + "href");
            if (!string.IsNullOrEmpty(href) && !href!.StartsWith("#", StringComparison.Ordinal) && !href.Contains("://") && !package.ContainsEntry(href)) {
                diagnostics.Add(new OdfDiagnostic("ODP102", OdfDiagnosticSeverity.Error,
                    $"Presentation image references missing package entry '{href}'.", "content.xml"));
            }
        }
    }
}
